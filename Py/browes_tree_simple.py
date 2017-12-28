import win32com.client
catia = win32com.client.Dispatch('catia.application')
productDocument1 = catia.ActiveDocument
part_list = []
products1 = catia.ActiveDocument.Product.Products
for part in range(1, products1.Count+1):
    part_list.append(products1.Item(part).PartNumber)
    print 'a'
print part_list

        
