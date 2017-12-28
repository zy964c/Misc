import json
seed = open('C:\Temp\zy964c\SEED_ALL.txt', 'r')
seed_data  =  list(seed)
seed_filtered = []
for n in seed_data:
    upd = n.replace(' ', '').replace('\n', '').replace(',', '').replace('"', '')
    seed_filtered.append(upd)

customer = open('C:\Temp\zy964c\AFA.txt', 'r')
x = json.load(customer)
#print x

parts_not_in_seed1 = []
for k in x:
    if k in seed_filtered:
        continue
        #print k + ' is in the part list'
    else:
        #print k + ' is NOT in the part list'
        if 'IR' not in k and 'CA' not in k:
            parts_not_in_seed1.append(k)

customer = open('C:\Temp\zy964c\OMR_and_KAL.txt', 'r')
x = json.load(customer)
#print x

parts_not_in_seed2 = []
for k in x:
    if k in seed_filtered:
        continue
        #print k + ' is in the part list'
    else:
        #print k + ' is NOT in the part list'
        if 'IR' not in k and 'CA' not in k:
            parts_not_in_seed2.append(k)
list_to_sort = []
unique_parts = set(parts_not_in_seed1 + parts_not_in_seed2)
for part in unique_parts:
    #print part
    list_to_sort.append(part)
list_to_sort.sort()
for n in list_to_sort:
    print n
    

        
