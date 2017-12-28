f = open('C:\Temp\zy964c\SEED_ALL.txt', 'r')
data  =  list(f)
#print data
filtered = []
for n in data:
    upd = n.replace(' ', '').replace('\n', '').replace(',', '').replace('"', '')
    filtered.append(upd)
print filtered
