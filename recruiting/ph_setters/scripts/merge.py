import csv,glob,json
base=list(csv.DictReader(open('/home/user/Claude_code/recruiting/ph_setters/candidates_base.csv')))
em={}
for p in glob.glob('/home/user/Claude_code/recruiting/ph_setters/emails/batch*.jsonl'):
    for line in open(p):
        line=line.strip()
        if not line: continue
        try: r=json.loads(line)
        except: continue
        if r.get('email'): em[r['profileUrl']]=r['email'].strip().lower()
out=[]; seen=set()
for r in base:
    e=em.get(r['profileUrl'])
    if e and e not in seen:
        seen.add(e); r['personal_email']=e; out.append(r)
cols=['firstname','lastname','personal_email','jobTitle','companyName','address','headline','profileUrl']
with open('/home/user/Claude_code/recruiting/ph_setters/ph_setter_candidates_with_emails.csv','w',newline='') as f:
    w=csv.DictWriter(f,fieldnames=cols,extrasaction='ignore'); w.writeheader(); w.writerows(out)
print('base',len(base),'looked up',len(em),'unique emails',len(out))
