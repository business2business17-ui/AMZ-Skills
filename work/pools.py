import re,json,collections
from kw import load,BAD,PROD
d=load()
def norm(k):
    k=k.lower()
    k=re.sub(r"\bi\s*phone\s*(?=\d)","iphone ",k); k=k.replace("ipone","iphone").replace("iphone","iphone ").replace("  "," ")
    k=re.sub(r"promax","pro max",k); k=re.sub(r"(\d{2})\s*(pro|plus|max|air|e\b)",r"\1 \2",k)
    k=re.sub(r"pro\s*max","pro max",k); k=re.sub(r"\s+"," ",k).strip()
    return k
FOREIGN=r"\bs\d{2}\b|\bz?\s?(fold|flip)\b|\bxr\b|\bxs\b|\bse\b|\bx\b|\biphone (?:[5-9]|1[0-2])\b|\b1[0-2] (?:pro|mini|plus)|\b1[0-8]e\b|\bair\b|\bultra\b|\bgalaxy|\bpixel|\bmini\b|\b1[3-8] mini\b|\bnote\b|\bs2\d|\bgeneration|\b(?:11|12) "
def mset(k):
    out=set()
    for m in re.finditer(r"\b(1[3-8])\b(?: (pro max|pro|plus))?",k):
        out.add((int(m.group(1)),m.group(2) or "base"))
    return out
COLOR=r"\b(black|white|pink|blue|red|green|purple|gold|silver|orange|yellow|gray|grey|brown|beige|teal|navy|rose|lavender|matte|frosted|colorful|glitter|marble|floral|flower|butterfly|cute|aesthetic|anime|cartoon|leather|silicone|liquid|wallet|card|kickstand|stand|ring|grip|rugged|heavy duty|military|waterproof|shockproof|drop|girls|women|men|kids|boys|baby|slim|thin|fit|tough|armor|defender|shell|hybrid|dual|layer|mag ?safe ?ring|bumper|screen|protector|lens|camera|crossbody|lanyard|strap|charging|wireless|support|compatible|qi2|qi 2|fashion|luxury|designer|bling|diamond|sparkle|stars?|heart|dog|cat|bear|sunflower|rainbow|hippie|boho|western|vintage|retro|gradient|case-mate|military grade)\b"
def target(g,v):
    return {(g,v)}|({(18,v)} if g==17 and v in("pro","pro max") else set())
MODELS=[(13,"base"),(13,"pro max"),(14,"base"),(14,"pro max"),(15,"base"),(15,"pro"),(15,"pro max"),(16,"base"),(16,"pro"),(16,"pro max"),(17,"base"),(17,"pro"),(17,"pro max")]
rows=[]
for k,v in d.items():
    if v['sv']<500 or v['ks']<100: continue
    n=norm(k)
    if re.search(BAD,n): continue
    rows.append((k,n,v))
pools={}
for g,var in MODELS:
    T=target(g,var); L=[]
    for k,n,v in rows:
        if re.search(FOREIGN,n): continue
        ms=mset(n)
        if ms and not ms<=T: continue
        if not ms:
            if not re.search(PROD,n): continue
        L.append({'kw':k,'sv':v['sv'],'ks':v['ks'],'model':bool(ms),'prod':bool(re.search(PROD,n)),'attr':bool(re.search(COLOR,n)),'files':sorted(v['files'])})
    pools[f"{g} {var}"]=L
json.dump(pools,open("pools_raw.json","w"),ensure_ascii=False)
for m,L in pools.items():
    print(m,len(L),"| with model:",sum(x['model'] for x in L),"| generic:",sum(not x['model'] for x in L),"| attr-free:",sum(not x['attr'] for x in L))
