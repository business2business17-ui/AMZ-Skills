import openpyxl,re,json,collections
BRAND="MOBILIUS"
SRC="/root/.claude/uploads/8a88f30f-2284-5dbc-ad73-f7ebfbaee185/bdf7c8b8-215_Amazon______________________11_described.xlsx"
def rows():
    w=openpyxl.load_workbook(SRC)["Sheet1"]
    return list(w.iter_rows(min_row=2,values_only=True))
def parse(r):
    model=r[5].strip().replace("Apple ","")
    print_name=r[4].split("/",1)[1].strip()
    d=r[3]
    m=re.match(r"Decorative phone case featuring (.+?) design\.\s*(.+?)\.\s*Built as",d)
    detail=m.group(2) if m else ""
    return model,print_name,detail
def models_txt(model):
    m18={"iPhone 17 Pro":"iPhone 18 Pro","iPhone 17 Pro Max":"iPhone 18 Pro Max"}
    return model,m18.get(model)
FILLER={"and","with","of","the","a","an","in","on","for","to"}
def short_print(pn,budget):
    words=[w for w in re.sub(r"[^A-Za-z0-9' -]"," ",pn).split() if w.lower() not in FILLER] or pn.split()
    out=""
    for w in words:
        cand=(out+" "+w).strip()
        if len(cand)<=budget: out=cand
        else: break
    return out or words[0]
def build(r,variant=0):
    model,pn,detail=parse(r)
    m17,m18=models_txt(model)
    if m18:
        mt=f"iPhone {m18.split()[1]}/{m17.split()[1]} {' '.join(m17.split()[2:])}"
        fits=f"{m18} and {m17}"; hl_fit=f"{m18}/{m17.replace('iPhone ','')}"
    else:
        mt=m17; fits=m17; hl_fit=m17
    suf=" Pattern" if variant else ""
    head=f"{BRAND} Clear Magnetic Case for {mt}"
    budget=75-len(head)-2-len(suf)
    ps=short_print(pn,budget)
    title=f"{head}, {ps}{suf}"
    noun="pattern" if variant else "print"
    hl=f"{pn} {noun} on a clear case; fits {hl_fit}; magnetic ring compatible with MagSafe"
    if len(hl)>125: hl=f"{pn} {noun} clear case; fits {hl_fit}; magnet ring compatible with MagSafe"
    d=(detail[0].upper()+detail[1:]) if detail else "Decorative print on the clear back"
    fit_txt=f"Precise fit for {fits} with a camera cutout that keeps the lens area open"
    b=[
     "MagSafe compatible magnetic ring: Built-in ring compatible with MagSafe chargers and accessories; N45 magnet, 2300 Gs, 55 mm outer and 46 mm inner diameter",
     f"Clear {m17} case: {fit_txt}" if m18 else f"Clear {m17} case: Precise fit with a camera cutout that keeps the lens area open",
     f"Clear TPU back: 2.12 mm thick TPU back panel lets your phone color show through the {pn.lower()} {noun}",
     f"{pn} {noun} design: {d}",
     "Slim, lightweight protection: Guards against everyday scratches and bumps while keeping the phone easy to hold",
    ]
    desc=(f"{BRAND} clear magnetic phone case with {pn} {noun} for {fits}.\n\n"
          f"{d}.\n\n"
          f"Does it work with MagSafe? The case has a built-in magnetic ring compatible with MagSafe chargers and accessories. The ring uses an N45 magnet with 2300 Gs magnetic strength; outer diameter 55 mm, inner diameter 46 mm, thickness 1.4 mm.\n\n"
          f"What is it made of? The back panel is 2.12 mm thick TPU, clear so your phone color shows through the {noun}. The camera cutout leaves the lens area open.\n\n"
          f"Slim, lightweight protection against everyday scratches and bumps.\n\n"
          f"Which phones does it fit? Compatible with {fits}.")
    return dict(title=title,hl=hl,bullets=b,desc=desc,model=model,pn=pn)
# Backend: only words whose phrases passed Search Volume >= 500 in Cerebro / Helium 10 (checked 2026-10-03)
NO_FORRO={"14","15 pro","16 pro"}          # "forro para iphone N" below 500 SV
def backend(model,o):
    key=model.replace("Apple iPhone ","").strip().lower()
    gen=key.split()[0]
    t=["transparent","funda","fundas","para","cover","cases"]
    if key not in NO_FORRO: t.insert(4,"forro")
    if key in("13","14","15","16","17"): t.append(f"iphone{gen}")
    if key in("16 pro","17 pro"): t.append(f"iphone{gen}pro")
    vis=" ".join([o['title'],o['hl'],*o['bullets'],o['desc']]).lower()
    vt=set(re.findall(r"[a-zñ0-9]+",vis))
    t=[w for w in t if w not in vt]
    return " ".join(t)
def checks(o):
    errs=[]
    t=o['title']
    if len(t)>75: errs.append(f"title {len(t)}>75")
    if re.search(r"[!$?_{}^¬¦]",t): errs.append("title special")
    words=re.findall(r"[A-Za-z0-9']+",t.lower()); c=collections.Counter(w for w in words if w not in{"for","with","and","the","of","a","in"})
    if any(v>2 for v in c.values()): errs.append("title repeat")
    if len(o['hl'])>125: errs.append(f"hl {len(o['hl'])}>125")
    tb=0
    for i,x in enumerate(o['bullets']):
        n=len(x); tb+=len(x.encode())
        if not 10<=n<=255: errs.append(f"b{i+1} len {n}")
        if x.endswith(".") or not x[0].isupper() or re.search(r"[™®©€…†‡¢£¥±~]|[^\x00-\x7f]",x): errs.append(f"b{i+1} format")
    if tb>1000: errs.append(f"bullets bytes {tb}>1000")
    if len(o['desc'])>2000: errs.append("desc>2000")
    return errs,tb
if __name__=="__main__":
    R=rows(); pick={}
    for r in R:
        mdl=r[5].strip()
        if mdl in("Apple iPhone 17 Pro Max","Apple iPhone 13") and len(pick.setdefault(mdl,[]))<3: pick[mdl].append(r)
    for mdl,L in pick.items():
        for r in L:
            o=build(r); e,tb=checks(o)
            print("\n=====",r[1],"|",mdl,"|",r[4]); print("TITLE",len(o['title']),o['title']); print("HL",len(o['hl']),o['hl'])
            for i,x in enumerate(o['bullets']): print(f"B{i+1} ({len(x)})",x)
            print("bullet bytes",tb,"| DESC",len(o['desc']),"| ERRORS:",e)
    allerr=collections.Counter(); titles=collections.Counter(); trunc=0
    for r in R:
        o=build(r); e,_=checks(o); titles[o['title']]+=1
        for x in e: allerr[re.sub(r"\d+","N",x)]+=1
        if o['pn'].lower()!=o['title'].split(", ")[1].lower(): trunc+=1
    print("\nALL",len(R),"errors:",dict(allerr),"| unique titles:",len(titles),"| titles with shortened print:",trunc)
    print("dup titles sample:",[t for t,c in titles.items() if c>1][:3])
