import openpyxl,glob,re,collections,json,warnings
warnings.filterwarnings("ignore")
U="/root/.claude/uploads/8a88f30f-2284-5dbc-ad73-f7ebfbaee185/"
def num(x):
    try: return float(str(x).replace(',',''))
    except: return None
def load():
    d={}
    for f in glob.glob(U+"*.xlsx"):
        if "described" in f: continue
        ws=openpyxl.load_workbook(f).active
        rows=list(ws.iter_rows(values_only=True)); h=rows[0]; i={n:k for k,n in enumerate(h)}
        if 'Search Volume' not in i: continue
        for r in rows[1:]:
            kw=str(r[0]).strip().lower(); sv=num(r[i['Search Volume']]); ks=num(r[i['Keyword Sales']])
            if sv is None or ks is None: continue
            e=d.setdefault(kw,{'sv':0,'ks':0,'files':set(),'cpr':None,'td':None})
            e['sv']=max(e['sv'],sv); e['ks']=max(e['ks'],ks); e['files'].add(f.split('/')[-1][9:])
            e['td']=r[i['Title Density']]
    return d
def models(kw):
    k=kw.lower().replace("promax","pro max").replace("i phone","iphone").replace("ipone","iphone")
    k=re.sub(r"(\d{2})(pro|plus|max|air)",r"\1 \2",k)
    out=set()
    for m in re.finditer(r"\b(1[0-9])\b(?:\s*(pro max|pro|plus|air|mini|e)\b)?",k):
        g=int(m.group(1)); v=m.group(2) or "base"
        out.add((g,v))
    return out
BAD=r"screen|protector|charger|charging|cable|wallet|holder|stand|airpods|watch|glass|lens|tempered|ipad|samsung|galaxy|pixel|case[- ]?mate|otterbox|spigen|casetify|speck|ugreen|esr|torras|mous|lifeproof|apple (?:store|care)|pop ?socket|popsocket|strap|lanyard|belt|clip|power bank|battery|adapter|usb|earbud|headphone|airtag|kickstand|ring light|tripod|gimbal|mount|vent|dock|stylus|pen|macbook|mac |ipad|android|moto|oneplus"
PROD=r"case|cases|cover|covers|funda|fundas|forro|cobertor|capa|estuche|carcasa|protective|bumper|shell"
if __name__=="__main__":
    d=load(); print(len(d))
    keep=[k for k,v in d.items() if v['sv']>=500 and v['ks']>=100]
    print("pass",len(keep))
    c=collections.Counter()
    for k in keep:
        if re.search(BAD,k): c['bad']+=1
        elif not re.search(PROD,k): c['noprod']+=1
        else: c['ok']+=1
    print(c)
    import random; random.seed(1)
    noprod=[k for k in keep if not re.search(BAD,k) and not re.search(PROD,k)]
    print(random.sample(noprod,30))
    ok=[k for k in keep if not re.search(BAD,k) and re.search(PROD,k)]
    gen=[k for k in ok if not models(k)]
    print("generic (no model):",len(gen)); print(sorted(gen,key=lambda k:-d[k]['ks'])[:60])
    json.dump({k:{'sv':d[k]['sv'],'ks':d[k]['ks'],'files':sorted(d[k]['files'])} for k in ok},open("ok.json","w"))
