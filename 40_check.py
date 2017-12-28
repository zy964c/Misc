import win32com.client
import json
catia = win32com.client.Dispatch('catia.application')
productDocument1 = catia.ActiveDocument
part_list = []
wrong_ids = []
def pn_outputer(prod_collection, parts):   
    for part in range(1, prod_collection.Count+1):
        try:
            parts.append(prod_collection.Item(part).Name)
            new_products_collection = prod_collection.Item(part).Products
            parts_new = pn_outputer(new_products_collection, parts)
        except:
            #print 'EXCEPTION OCCURED'
            continue
    return parts
products1 = catia.ActiveDocument.Product.Products
max_symbols_allowed = raw_input('Enter MAX symbols allowed:\n')
part_list1 = pn_outputer(products1, part_list)
#print part_list1
for item in part_list1:
    #print len(str(item))
    if len(str(item)) > int (max_symbols_allowed):
        wrong_ids.append(str(item))
if len(wrong_ids) > 0:
    print "Here is the list of instances with more than " + max_symbols_allowed + " symbols in instance ID:\n"
    for elem in wrong_ids:
        print elem + " - " + str(len(str(elem))) + " symbols"
else:
    print "No errors found"



        
