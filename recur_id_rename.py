import win32com.client

catia = win32com.client.Dispatch('catia.application')
def pn_outputer(prod_collection, parts):   
    for part in range(1, prod_collection.Count+1):
        try:
            parts.append(prod_collection.Item(part).Name)
            cur_name = prod_collection.Item(part).Name
            new_name = cur_name.replace('.1', '_1')
            prod_collection.Item(part).Name = new_name
            if cur_name != new_name:
                print prod_collection.Item(part).Name
            new_products_collection = prod_collection.Item(part).ReferenceProduct.Products
            parts_new = pn_outputer(new_products_collection, parts)
        except:
            print 'EXCEPTION OCCURED'
            continue
    return parts

products1 = catia.ActiveDocument.Product.Products
part_list = []
part_list1 = pn_outputer(products1, part_list)




        
