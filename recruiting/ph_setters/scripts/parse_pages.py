import json,glob,csv,sys
rows={}
def load(p):
    t=open(p).read(); d=json.loads(t)
    while isinstance(d,list): d=json.loads(d[0]['text'])
    return d
for p in sorted(glob.glob('/home/user/Claude_code/recruiting/ph_setters/pages/page*.txt')):
    d=load(p)
    for l in d['leads']:
        u=l.get('profileUrl')
        if u and u not in rows: rows[u]=l
    last=d.get('next_cursor')
cols=['firstname','lastname','jobTitle','companyName','companyWebsite','address','headline','profileUrl']
with open('/home/user/Claude_code/recruiting/ph_setters/candidates_base.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(cols)
    for l in rows.values(): w.writerow([(l.get(c) or '').replace('\n',' ') for c in cols])
print(len(rows),'unique; last cursor:',last)
