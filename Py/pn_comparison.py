f = open('C:\Temp\zy964c\KAL.txt', 'r')
data  =  list(f)
filtered = []
for n in data:
    upd = n.replace(' ', '').replace('\n', '').replace(',', '').replace('"', '')
    filtered.append(upd)
print filtered
