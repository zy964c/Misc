import json
seed = open('C:\Temp\zy964c\AAA.txt', 'r')
seed_data  =  list(seed)
seed_filtered = []
for n in seed_data:
    upd = n.replace(' ', '').replace('\n', '').replace(',', '').replace('"', '')
    seed_filtered.append(upd)

seed_filtered.sort()
for n in seed_filtered:
    print n
