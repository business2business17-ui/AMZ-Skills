import openpyxl,re,subprocess,sys
R="/home/user/business2business17-ui/"
w=openpyxl.load_workbook(R+"223/223 Amazon New 17.09 - with prints.xlsx").active
p={}
for r in w.iter_rows(min_row=2,values_only=True):
    m=re.match(r"223(\w+?)-(F\d+)Ru001q(\d+)",r[0]); p[(m.group(2),m.group(3))]=(r[11],r[12],r[13])
pdf=R+"223/223 Amazon (1)-1.pdf" if sys.argv[1]=="full" else R+"223/catalog_223_u001q001_u001q110_2026-09-17_10-41-03-1.pdf"
for pg in map(int,sys.argv[2:]):
    t=subprocess.run(["pdftotext","-f",str(pg),"-l",str(pg),"-layout",pdf,"-"],capture_output=True,text=True).stdout
    labs=re.findall(r"223xxx-(F\d+)Ru001q(\d+)",t); print(f"\n=== {sys.argv[1]} page {pg}")
    for i in range(0,len(labs),5):
        print(" | ".join(f"{f}q{q}:"+(p[(f,q)][0]+" ["+p[(f,q)][2]+"]" if (f,q) in p else "-") for f,q in labs[i:i+5]))
