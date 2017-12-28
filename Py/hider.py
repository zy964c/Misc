import win32com.client
import json
catia = win32com.client.Dispatch('catia.application')
productDocument1 = catia.ActiveDocument
productSelection1 = catia.ActiveDocument.Selection
productSelection2 = catia.ActiveDocument.Selection
part_list = []
def hider(prod_collection, name):  
    for part in range(1, prod_collection.Count+1):
        try:
            if prod_collection.Item(part).Name != name:
                productSelection2.add(prod_collection.Item(part))
                new_products_collection = prod_collection.Item(part).Products
                parts_new = hider(new_products_collection, selected1_name)
            else:
                continue
        except:
            #print 'EXCEPTION OCCURED'
            continue    
    productSelection2.visProperties.SetShow(1)
    productSelection2.Clear

products1 = catia.ActiveDocument.Product.Products
print 'Select part to keep in show'
productSelection1.SelectElement3(['AnyObject'],'Select part to keep in show', False, 0, False)
selected1 = productSelection1.Item2(1).Value
selected1_name = selected1.Name
print selected1_name
productSelection1.Clear
hider(products1, selected1_name)
