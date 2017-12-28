import win32com.client
from Tkinter import *

catia = win32com.client.Dispatch('catia.application')
productDocument1 = catia.ActiveDocument
part_list = []
wrong_ids = []

class App:

    def __init__(self, master):

        frame = Frame(master)
        frame.pack()

        i = IntVar()
        self.lable = Label(frame, text="Enter MAX symbols allowed")
        self.lable.pack
        self.entry = Entry(frame, textvariable = i, width = 10)
        self.entry.pack

#root = Tk()

#app = App(root)


def pn_outputer(prod_collection, parts):   
    for part in range(1, prod_collection.Count+1):
        try:
            name_new = ""
            name = prod_collection.Item(part).Name         
            #if " " in name:
                #for i in name:
                    #if i != " ":
                        #name_new += i
                #prod_collection.Item(part).Name = name_new
                #prod_collection.Item(part).Update
                #print "!!!\n" + name_new
            parts.append(prod_collection.Item(part).Name)
            new_products_collection = prod_collection.Item(part).Products
            parts_new = pn_outputer(new_products_collection, parts)
        except:
            continue
    return parts
products1 = catia.ActiveDocument.Product.Products
max_symbols_allowed = raw_input('Enter MAX symbols allowed:\n')
part_list1 = pn_outputer(products1, part_list)
for item in part_list1:
    if len(str(item)) > int (max_symbols_allowed):
        wrong_ids.append(str(item))
if len(wrong_ids) > 0:
    print "Here is the list of instances with more than " + max_symbols_allowed + " symbols in instance ID:\n"
    for elem in wrong_ids:
        print elem + " - " + str(len(str(elem))) + " symbols"
else:
    print "No errors found"



#root.mainloop()

        
