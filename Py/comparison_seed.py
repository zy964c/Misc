seed = open('C:\Temp\zy964c\SEED_ALL.txt', 'r')
seed_data  =  list(seed)
seed_filtered = []
for n in seed_data:
    upd = n.replace(' ', '').replace('\n', '').replace(',', '').replace('"', '')
    seed_filtered.append(upd)

f = open('C:\Temp\zy964c\KLM.txt', 'r')
data  =  list(f)
filtered = []
for n in data:
    upd = n.replace(' ', '').replace('\n', '').replace(',', '').replace('"', '')
    filtered.append(upd)

parts_not_in_seed = []
for k in filtered:
    if k in seed_filtered:
        print k + ' is in the part list'
    else:
        print k + ' is NOT in the part list'
        parts_not_in_seed.append(k)

print 'Here is list of parts without SEED model: '
for i in parts_not_in_seed:
    print i

        
