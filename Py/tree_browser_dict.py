import win32com.client
import json
catia = win32com.client.Dispatch('catia.application')
productDocument1 = catia.ActiveDocument
part_list = []
instace_IDs = []
def pn_outputer(prod_collection, parts, instances):   
    for part in range(1, prod_collection.Count+1):
        try:
            instances.append(prod_collection.Item(part).Name)
            parts.append(prod_collection.Item(part).PartNumber)
            new_products_collection = prod_collection.Item(part).Products
            parts_new = pn_outputer(new_products_collection, parts, instances)
        except:
            print 'EXCEPTION OCCURED'
            continue
    dict1 = dict(zip(parts, instances))
    return dict1

print 'gg'
products1 = catia.ActiveDocument.Product.Products
part_list1 = pn_outputer(products1, part_list, instace_IDs)
print part_list1
        
