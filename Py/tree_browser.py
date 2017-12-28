import win32com.client
import json
catia = win32com.client.Dispatch('catia.application')
productDocument1 = catia.ActiveDocument
part_list = []
def pn_outputer(prod_collection, parts):   
    for part in range(1, prod_collection.Count+1):
        try:
            parts.append(prod_collection.Item(part).PartNumber)
            new_products_collection = prod_collection.Item(part).Products
            parts_new = pn_outputer(new_products_collection, parts)
        except:
            print 'EXCEPTION OCCURED'
            continue
    return parts
products1 = catia.ActiveDocument.Product.Products
part_list1 = pn_outputer(products1, part_list)
part_set = set(part_list1)
output_val = []
for item in part_set:
    output_val.append(item)
file_name = raw_input('Please enter file name:')
path = 'C:\\Temp\\zy964c\\'
with open(path + str(file_name) + '.txt' , 'w') as f:
    json.dump(output_val, f)
print 'file saved to: ' + path + file_name + '.txt'



        
