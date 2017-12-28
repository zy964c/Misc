import win32com.client
catia = win32com.client.Dispatch('catia.application')
productDocument1 = catia.ActiveDocument
part_list = []
products1 = catia.ActiveDocument.Product.Products
for part in range(1, products1.Count):
    part_list.append(products1.Item(part).PartNumber)

f = open('C:\Temp\zy964c\KAL.txt', 'r')
data  =  list(f)
filtered = []
for n in data:
    upd = n.replace(' ', '').replace('\n', '').replace(',', '').replace('"', '')
    filtered.append(upd)

parts_not_in_seed = []
for k in filtered:
    if k in part_list:
        print k + ' is in part list'
    else:
        print k + ' is not in part list'
        parts_not_in_seed.append(k)

print 'here are parts without SEED model: '
for i in parts_not_in_seed:
    print i

        
