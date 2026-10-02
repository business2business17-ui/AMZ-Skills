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
def build(r):
    model,pn,detail=parse(r)
    m17,m18=models_txt(model)
    # title: brand + type + model(+18) + attr
    if m18:
        title=f"{BRAND} Clear Magnetic Case for {m17}/{m18.replace('iPhone ','')}, MagSafe Compatible"
    else:
        title=f"{BRAND} Clear Magnetic Case for {m17}, Compatible with MagSafe"
    fits=f"{m17} and {m18}" if m18 else m17
    hl=f"{pn} print on a clear case for {m17}; magnetic ring compatible with MagSafe"
    b=[
     "Magnetic ring: Built-in ring compatible with MagSafe chargers and accessories; N45 magnet, 2300 Gs, 55 mm outer and 46 mm inner diameter",
     f"Fits {fits}: Precise camera cutout keeps the lens area open; clear back lets the phone color show through",
     f"Clear TPU back: 2.12 mm thick TPU back panel in a slim, lightweight case that guards against everyday scratches and bumps",
     f"{pn} design: {detail[0].upper()+detail[1:]}" if detail else f"{pn} design on the clear back",
     f"Phone case with print: Decorative {pn.lower()} print for {m17}, easy to pair with MagSafe chargers and accessories",
    ]
    desc=(f"{BRAND} clear phone case with {pn} design for {fits}.\n\n"
          f"{detail}.\n\n"
          f"The case has a built-in magnetic ring compatible with MagSafe chargers and accessories. The ring uses an N45 magnet with 2300 Gs magnetic strength; outer diameter 55 mm, inner diameter 46 mm, thickness 1.4 mm. "
          f"The back panel is 2.12 mm thick TPU, and the camera cutout leaves the lens area open. The clear back lets your phone color show through the print.\n\n"
          f"Compatible with {fits}.")
    return dict(title=title,hl=hl,bullets=b,desc=desc,model=model,pn=pn)
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
        if mdl in("Apple iPhone 17 Pro Max","Apple iPhone 13","Apple iPhone 14") and mdl not in pick: pick[mdl]=r
    for mdl,r in pick.items():
        o=build(r); e,tb=checks(o)
        print("\n=====",r[1],"|",mdl)
        print("TITLE",len(o['title']),o['title']); print("HL",len(o['hl']),o['hl'])
        for i,x in enumerate(o['bullets']): print(f"B{i+1} ({len(x)})",x)
        print("bullet bytes",tb,"| DESC",len(o['desc'])); print(o['desc']); print("ERRORS:",e)
