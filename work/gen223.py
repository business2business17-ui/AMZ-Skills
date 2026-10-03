import openpyxl,re,collections,json
BRAND="MOBILIUS"
SRC="/root/.claude/uploads/8a88f30f-2284-5dbc-ad73-f7ebfbaee185/b3172287-223_Amazon_New_17.09_described.xlsx"
FILLER={"and","with","of","the","a","an","in","on","for","to"}
def rows():
    ws=openpyxl.load_workbook(SRC).active; out=[]; seen=set()
    for r in ws.iter_rows(min_row=2,values_only=True):
        if r[1] in seen: continue            # exact duplicate rows removed
        seen.add(r[1]); out.append(r)
    return out
def short_print(pn,budget):
    words=[w for w in re.sub(r"[^A-Za-z0-9' -]"," ",pn).split() if w.lower() not in FILLER] or pn.split()
    out=""
    for w in words:
        cand=(out+" "+w).strip()
        if len(cand)<=budget: out=cand
        else: break
    return out or words[0]
MODEL={ # display in title, fits text, key for backend
 "Apple iPhone 16e/Apple iPhone 17e":("iPhone 16e/17e","iPhone 16e and iPhone 17e","16e17e"),
 "Apple iPhone 17 Pro / 18 Pro":("iPhone 18/17 Pro","iPhone 18 Pro and iPhone 17 Pro","17 pro"),
 "Apple iPhone 17 Pro Max / 18 Pro Max":("iPhone 18/17 Pro Max","iPhone 18 Pro Max and iPhone 17 Pro Max","17 pro max"),
}
def model_info(m):
    m=m.strip()
    if m in MODEL: return MODEL[m]
    n=m.replace("Apple ",""); return (n,n,n.replace("iPhone ","").lower())
FIN={
 "Black Smoky":dict(back="Smoky matte back",txt="2.12 mm thick TPU back panel with a translucent smoky dark grey matte finish, black frame and black buttons",hl="black smoky matte"),
 "Black Frosted":dict(back="Frosted matte back",txt="2.12 mm thick TPU back panel with a translucent light frosted grey finish, black frame and black buttons",hl="black frosted matte"),
 "White Frosted":dict(back="Frosted matte back",txt="2.12 mm thick TPU back panel with a translucent frosted clear finish, white frame and white buttons",hl="white frosted matte"),
 "Orange":dict(back="Solid orange finish",txt="Opaque soft-touch orange case with an orange camera surround and orange buttons; 2.12 mm thick TPU back panel",hl="orange matte"),
}
def parse(r):
    col,pn=[x.strip() for x in r[4].split("/",1)]
    m=re.match(r"Decorative phone case featuring (.+?) design\.\s*(.+?)\.\s*Built as",r[3])
    detail=m.group(2) if m else ""
    return col,pn,detail
def _build(r,variant=0,tier="A"):
    col,pn,detail=parse(r); mt,fits,_=model_info(r[5]); suf=" Pattern" if variant else ""
    mid="Matte Magnetic Case" if tier=="A" else "Matte Case"
    head=f"{BRAND} {col} {mid} for {mt}"
    ps=short_print(pn,75-len(head)-2-len(suf)); title=f"{head}, {ps}{suf}"
    noun="pattern" if variant else "print"
    f=FIN[col]
    art="an" if f["hl"][0] in "aeiou" else "a"
    hl=f"{pn} {noun} on {art} {f['hl']} case; fits {mt}; magnetic ring compatible with MagSafe"
    if len(hl)>125: hl=f"{pn} {noun}, {f['hl']} case; fits {mt}; magnet ring compatible with MagSafe"
    if len(hl)>125: hl=f"{short_print(pn,24)} {noun}, {f['hl']} case; fits {mt}; magnet ring compatible with MagSafe"
    d=(detail[0].upper()+detail[1:]) if detail else "Decorative print on the case back"
    single=(fits==mt)
    b=[
     "MagSafe compatible magnetic ring: Built-in ring compatible with MagSafe chargers and accessories; N45 magnet, 2300 Gs, 55 mm outer and 46 mm inner diameter",
     f"Matte {mt} case: Precise fit with a camera cutout that keeps the lens area open" if single else f"Matte {mt} case: Precise fit for {fits} with a camera cutout that keeps the lens area open",
     f"{f['back']}: {f['txt'] if col!='Orange' else f['txt']}",
     f"{pn} {noun} design: {d}",
     "Slim, lightweight protection: Guards against everyday scratches and bumps while keeping the phone easy to hold",
    ]
    desc=(f"{BRAND} {col.lower()} matte magnetic phone case with {pn} {noun} for {fits}.\n\n{d}.\n\n"
          f"Does it work with MagSafe? The case has a built-in magnetic ring compatible with MagSafe chargers and accessories. The ring uses an N45 magnet with 2300 Gs magnetic strength; outer diameter 55 mm, inner diameter 46 mm, thickness 1.4 mm.\n\n"
          f"What is it made of and how does it look? {f['txt'][0].upper()+f['txt'][1:]}. The camera cutout leaves the lens area open.\n\n"
          f"Slim, lightweight protection against everyday scratches and bumps.\n\nWhich phones does it fit? Compatible with {fits}.")
    return dict(title=title,hl=hl,bullets=b,desc=desc,col=col,pn=pn,model=r[5].strip())

_US=[(r"watercolour","watercolor"),(r"colourful","colorful"),(r"colours","colors"),(r"colour","color"),(r"greys","grays"),(r"grey","gray"),(r"centred","centered"),(r"centres","centers"),(r"centre","center"),(r"cosy","cozy")]
def us(t):
    for a,b in _US:
        t=re.sub(a,b,t); t=re.sub(a.capitalize(),b.capitalize(),t)
    return t
def us_all(o):
    o=dict(o); o['title']=us(o['title']); o['hl']=us(o['hl']); o['bullets']=[us(x) for x in o['bullets']]; o['desc']=us(o['desc']); o['pn']=us(o['pn']); return o
def build(r,variant=0,tier="A"): return us_all(_build(r,variant,tier))
# backend: only words from phrases with Search Volume >= 500 (Helium 10, 2026-10-03)
BK={"13":"funda fundas para forro cover cases iphone13","13 pro max":"funda fundas para forro cover cases","14":"funda fundas para cover cases iphone14","14 pro max":"funda fundas para forro cover cases",
 "15":"funda fundas para forro cover cases iphone15","15 pro":"funda fundas para cover cases","15 pro max":"funda fundas para forro cover cases","16":"funda fundas para forro cover cases iphone16",
 "16 plus":"funda fundas para cover cases","16 pro":"funda fundas para cover cases iphone16pro","16 pro max":"funda fundas para forro cover cases","16e17e":"funda fundas para cover cases iphone16e iphone17e",
 "17":"funda fundas para forro cover cases iphone17","17 pro":"funda fundas para forro cover cases iphone17pro","17 pro max":"funda fundas para forro cover cases"}
def backend(r,o):
    key=model_info(r[5])[2]; vis=set(re.findall(r"[a-zñ0-9]+"," ".join([o['title'],o['hl'],*o['bullets'],o['desc']]).lower()))
    return " ".join(w for w in BK[key].split() if w not in vis)
def checks(o):
    e=[]; t=o['title']
    if len(t)>75: e.append("title>75")
    if re.search(r"[!$?_{}^¬¦]",t): e.append("title special")
    wc=collections.Counter(w for w in re.findall(r"[a-z0-9']+",t.lower()) if w not in{"for","with","and","the","of","a","in"})
    if wc and max(wc.values())>2: e.append("title repeat")
    if len(o['hl'])>125: e.append("hl>125")
    tb=0
    for i,x in enumerate(o['bullets']):
        tb+=len(x.encode())
        if not 10<=len(x)<=255: e.append(f"b len")
        if x.endswith(".") or not x[0].isupper() or re.search(r"[^\x00-\x7f]",x): e.append("b format")
    if tb>1000: e.append("bullets>1000B")
    if len(o['desc'])>2000: e.append("desc>2000")
    return e,tb
if __name__=="__main__":
    R=rows(); print(len(R),"unique SKU rows")
    for tier in "AB":
        full=0;ts=collections.Counter();over=0
        for r in R:
            o=build(r,0,tier); ts[o['title']]+=1; over+=len(o['title'])>75
            full+= short_print(o['pn'],999).lower()==o['title'].split(", ")[-1].lower()
        print(tier,"full print:",full,"/",len(R),"| unique titles",len(ts),"| >75:",over)
