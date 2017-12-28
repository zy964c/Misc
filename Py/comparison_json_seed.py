import json
seed = open('C:\Temp\zy964c\SEED_ALL.txt', 'r')
seed_data  =  list(seed)
seed_filtered = []
for n in seed_data:
    upd = n.replace(' ', '').replace('\n', '').replace(',', '').replace('"', '')
    seed_filtered.append(upd)

customer = open('C:\Temp\zy964c\BRI.txt', 'r')
x = json.load(customer)
print x

parts_not_in_seed = []
for k in x:
    if k in seed_filtered:
        print k + ' is in the part list'
    else:
        print k + ' is NOT in the part list'
        if 'IR' not in k and 'CA' not in k:
            parts_not_in_seed.append(k)

print 'Here is list of parts without SEED model: '
for i in parts_not_in_seed:
    print i

        
