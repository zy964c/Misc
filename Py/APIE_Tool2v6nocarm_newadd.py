import Tkinter
import Tkconstants
import tkFileDialog
import tkMessageBox
import math
import time

import win32com.client

from carm1 import CARM
from carm import ProductECS

CATIA = win32com.client.Dispatch('catia.application')
ICM = CATIA.ActiveDocument
oFileSys = CATIA.FileSystem

ICM_1 = ICM.Product
ICM_Products = ICM_1.Products

global angle
angle = 0

global itemProd
itemProd = 5

global bin_breaker
bin_breaker = []

global order_of_templete_product
order_of_templete_product = 4

global sta_value_pairs
sta_value_pairs = []

global sta_values_fake
sta_values_fake = []

global make_carms
make_carms = False

global dash_number
dash_number = 1000

class TkFileDialogExample(Tkinter.Frame):
    def __init__(self, root):

        Tkinter.Frame.__init__(self, root)

        #f = StringVar()
        self.plug = Tkinter.IntVar()
        root.geometry("350x280")
        self.repl = Tkinter.IntVar()

        # options for buttons
        button_opt = {'fill': Tkconstants.BOTH, 'padx': 5, 'pady': 5}

        # define buttons

        self.sp = Tkinter.StringVar()
        self.lp = Tkinter.StringVar()

        #v.set("a default value")
        #s = v.get()
        Tkinter.Button(self, text='Choose a bin run', command=self.askopenfile).pack(**button_opt)
        Tkinter.Radiobutton(root, text="-8", variable=self.plug, value=0).pack(**button_opt)
        Tkinter.Radiobutton(root, text="-9", variable=self.plug, value=240).pack(**button_opt)
        Tkinter.Radiobutton(root, text="-10", variable=self.plug, value=456).pack(**button_opt)
        Tkinter.Label(root, text="Enter library path:").pack()
        Tkinter.Entry(root, textvariable=self.lp, width=50).pack()
        Tkinter.Label(root, text="Enter SOW path:").pack()
        Tkinter.Entry(root, textvariable=self.sp, width=50).pack()
        Tkinter.Label(root, text="Make CARMs").pack()
        Tkinter.Checkbutton(root, variable=self.repl).pack()

        self.plug.set(240)
        self.sp.set("\\\\FIL-MOW01-01\\787Payloads\\IRC\\SYSTEM_Int\\APIE_TOOL\\APIE_Tool2\\HNA_OMF_SOW.txt")
        self.lp.set("\\\\FIL-MOW01-01\\787Payloads\\IRC\SYSTEM_Int\\APIE_TOOL\\APIE_Tool2\\LIBRARY_NOGEOM_ICM2")

        # define options for opening or saving a file
        self.file_opt = options = {}
        options['defaultextension'] = '.txt'
        options['filetypes'] = [('all files', '.*'), ('text files', '.txt')]
        options['initialdir'] = 'C:\\'
        options['initialfile'] = 'myfile.txt'
        options['parent'] = root
        options['title'] = 'This is a title'

    def askopenfile(self):

        """Returns an opened file in read mode."""

        global make_carms

        f = tkFileDialog.askopenfile(parent=root, mode='rb', title='Choose a file')

        s1 = f.readline()
        if s1.startswith("#"):
            s1 = f.readline().replace(' ', '').replace('fairing', '1').replace('prem', '2').replace('EXT', '3').replace(
                '\r\n', '').split(",")
        s2 = f.readline()
        if s2.startswith("#"):
            s2 = f.readline().replace(' ', '').replace('fairing', '1').replace('prem', '2').replace('EXT', '3').replace(
                '\r\n', '').split(",")
        s3 = f.readline()
        if s3.startswith("#"):
            s3 = f.readline().replace(' ', '').replace('fairing', '1').replace('prem', '2').replace('EXT', '3').replace(
                '\r\n', '').split(",")
        s4 = f.readline()
        if s4.startswith("#"):
            s4 = f.readline().replace(' ', '').replace('fairing', '1').replace('prem', '2').replace('EXT', '3').replace(
                '\r\n', '').split(",")
        s5 = f.readline()
        if s5.startswith("#"):
            s5 = f.readline().replace(' ', '').replace('fairing', '1').replace('prem', '2').replace('EXT', '3').replace(
                '\r\n', '').split(",")
        s6 = f.readline()
        if s6.startswith("#"):
            s6 = f.readline().replace(' ', '').replace('fairing', '1').replace('prem', '2').replace('EXT', '3').replace(
                '\r\n', '').split(",")

        s1 = s1[::-1]
        s2 = s2[::-1]

        print s1  # just for checking list correctness
        print s2  # just for checking list correctness
        print s3  # just for checking list correctness
        print s4  # just for checking list correctness
        print s5  # just for checking list correctness
        print s6  # just for checking list correctness

        if self.plug.get() == 0:
            plug = 0
        elif self.plug.get() == 456:
            plug = 456
        else:
            plug = 240

        global path
        path = str(self.lp.get())
        print path
        global SOW
        SOW = str(self.sp.get())
        print SOW

        nonconst_comps_inst()

        if s3 != ['']:
            AddComponent(s3, 'LH', 'constant', 'middle', plug)
        if s4 != ['']:
            AddComponent(s4, 'RH', 'constant', 'middle', plug)
        if s1 != ['']:
            AddComponent(s1, 'LH', 'nonconstant', 'nose', 0)
        if s2 != ['']:
            AddComponent(s2, 'RH', 'nonconstant', 'nose', 0)
        if s5 != ['']:
            AddComponent(s5, 'LH', 'nonconstant', 'tail', plug)
        if s6 != ['']:
            AddComponent(s6, 'RH', 'nonconstant', 'tail', plug)

        global constSize
        sizes = ['18', '24', '30', '36', '42', '48', '54', '60', '72']
        constant_lenght = []
        constant_items = [s3, s4]
        for item in constant_items:
            for i in item:
                if i[:2] in sizes:
                    constant_lenght.append(i)
                else:
                    continue
        print constant_lenght        
        constSize = len(constant_lenght)

        global order_of_new_product
        order_of_new_product = dash_number - 995 + constSize

        if self.repl.get() == 1:
            #AddComponent(s3, 'LH', 'constant', 'middle', plug)
            make_carms = True

        #part_number_collector()

        #raw_input('Please bring to ENOVIA listed part references and press ENTER')

        #print 'it worked'

        #Replacer()

        #AddNewIRMs()

        ReframeAll()

        IRMer_nonc()

        IRMer_const()

        Deleter()

        tkMessageBox.showinfo(title="APIE", message="Done")
        root.withdraw()
        root.destroy()

def Inch_to_mm(distance):
    return distance * 25.4

def mm_to_Inch(distance):
    return distance / 25.4

def STAvalue(coord, plug_value):
    if plug_value == 240:
        if coord <= round(Inch_to_mm(609)):
            STA = '0' + str(int(round(coord / 25.4)))
        elif coord > round(Inch_to_mm(609)) and coord <= round(Inch_to_mm(609 + 120)):
            STA = '0609+' + str(int(round(coord / 25.4 - 609)))
        elif coord > round(Inch_to_mm(609 + 120)) and coord <= round(Inch_to_mm(1401 + 120)):
            if (coord / 25.4 - 120) < 1000:
                STA = '0' + str(int(round(coord / 25.4 - 120)))
            else:
                STA = str(int(round(coord / 25.4 - 120)))
        elif coord > round(Inch_to_mm(1401 + 120)) and coord <= round(Inch_to_mm((1401 + 120) + 120)):
            STA = '1401+' + str(int(round(coord / 25.4 - (1401 + 120))))
        elif coord > round(Inch_to_mm(1401 + 240)):
            STA = str(int(round(coord / 25.4 - 240)))

    elif plug_value == 456:
        if coord <= round(Inch_to_mm(609)):
            STA = '0' + str(int(round(coord / 25.4)))
        elif coord > round(Inch_to_mm(609)) and coord <= round(Inch_to_mm(609 + 240)):
            STA = '0609+' + str(int(round(coord / 25.4 - 609)))
        elif coord > round(Inch_to_mm(609 + 240)) and coord <= round(Inch_to_mm(1401 + 240)):
            if (coord / 25.4 - 240) < 1000:
                STA = '0' + str(int(round(coord / 25.4 - 240)))
            else:
                STA = str(int(round(coord / 25.4 - 240)))
        elif coord > round(Inch_to_mm(1401 + 240)) and coord <= round(Inch_to_mm((1401 + 240) + 120)):
            STA = '1401+' + str(int(round(coord / 25.4 - (1401 + 240))))
        elif coord > round(Inch_to_mm(1401 + 360)) and coord <= round(Inch_to_mm(1618 + 360)):
            STA = str(int(round(coord / 25.4 - 360)))
        elif coord > round(Inch_to_mm(1618 + 360)) and coord <= round(Inch_to_mm((1618 + 360) + 96)):
            STA = '1618+' + str(int(round(coord / 25.4 - (1618 + 360))))
        elif coord > round(Inch_to_mm(1618 + 360 + 96)):
            STA = str(int(round(coord / 25.4 - (360 + 96))))

    elif plug_value == 0:
        if int(round(coord / 25.4)) < 1000:
            STA = '0' + str(int(round(coord / 25.4)))
        else:
            STA = str(int(round(coord / 25.4)))
    return STA

def selector(sel1, prod1, lower=False):
    global constSize
    global order_of_new_product
    sel1.Copy()
    sel1.Clear()
    productDocument1 = CATIA.ActiveDocument
    selection2 = productDocument1.Selection
    selection2.Clear()
    print (itemProd + constSize)
    order_of_new_product = (itemProd + constSize)
    if not lower:
        product_forpaste = prod1.Item(order_of_new_product)
    else:
        product_forpaste = prod1.Item(order_of_new_product+1)
    print product_forpaste.PartNumber
    selection2.add(product_forpaste)
    selection2.Paste()
    selection2.Clear()

def Replacer():
    ICM_1.ApplyWorkMode(2)
    replacedDetails = []
    revC = ["832Z4501-1"]
    revB = ["832Z4501-10", "832Z4501-3", "832Z4501-4", "832Z4501-5", "832Z4501-7", "832Z4501-6", "832Z4501-2",
            "832Z4501-8", "832Z4501-9", "832Z4501-11"]
    revA = ["830Z1009-2724", "830Z1009-2736", "830Z1009-2748"]
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products

    for prod in xrange(1, 5):
        product_to_replace = products1.Item(prod)
        products_to_replace = product_to_replace.Products
        for det in xrange(1, products_to_replace.Count+1):
            product_act_to_replace_nonc = products_to_replace.Item(det)
            products_act_to_replace_nonc = product_act_to_replace_nonc.Products
            for det_deep in xrange(1, products_act_to_replace_nonc.Count):
                product_act_to_replace = products_act_to_replace_nonc.Item(det_deep)
                if "CA" in str(product_act_to_replace.PartNumber):
                    continue
                else:
                    y = product_act_to_replace.PartNumber
                    replacedDetail = str(y)
                    if replacedDetail in replacedDetails:
                        continue
                    else:
                        replacedDetails.append(replacedDetail)
                        replaced = str(y)
                        print replacedDetail
                        if replaced in revB:
                            replacing = replaced + '--B.CATPart'
                        elif replaced in revA:
                            replacing = replaced + '--A.CATPart'
                        elif replaced in revC:
                            replacing = replaced + '--C.CATPart'
                        else:
                            replacing = replaced + '.CATPart'
                        print replacing

                        documents1 = CATIA.Documents
                        partDocument1 = documents1.Item(replacing)
                        Prod_replacing_part = partDocument1.GetItem(replaced)
                        products_act_to_replace_nonc.ReplaceProduct(product_act_to_replace,
                                                                               Prod_replacing_part, True)

    for prod in xrange(5, products1.Count+1):
        product_to_replace = products1.Item(prod)
        products_to_replace = product_to_replace.Products

        for det in xrange(1, products_to_replace.Count):
            product_act_to_replace = products_to_replace.Item(det)
            if "CA" in str(product_act_to_replace.PartNumber):
                continue
            else:
                y = product_act_to_replace.PartNumber
                replacedDetail = str(y)
                if replacedDetail in replacedDetails:
                    continue
                else:
                    replacedDetails.append(replacedDetail)
                    replaced = str(y)
                    print replacedDetail
                    if replaced in revB:
                        replacing = replaced + '--B.CATPart'
                    elif replaced in revA:
                        replacing = replaced + '--A.CATPart'
                    else:
                        replacing = replaced + '.CATPart'
                    print replacing

                    documents1 = CATIA.Documents
                    partDocument1 = documents1.Item(replacing)
                    Prod_replacing_part = partDocument1.GetItem(replaced)
                    products_to_replace.ReplaceProduct(product_act_to_replace, Prod_replacing_part, True)

def AddNewIRMs():
    alldata = PartNumberCreator(SOW)
    pn = alldata[0]
    print pn
    id = alldata[1]
    print id
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products

    for item in range(len(pn)):
        product2 = products1.AddNewComponent("Product", pn[item])
        product2.name = id[item]

def ReframeAll():

    ICM_1.ApplyWorkMode(2)
    specsAndGeomWindow1 = CATIA.ActiveWindow
    viewer3D1 = specsAndGeomWindow1.ActiveViewer
    viewer3D1.Reframe()
    viewer3D1.Viewpoint3D

def PartNumberCreator(path1):
    with open(path1, 'r') as f:
        read_data = f.readlines()
    print read_data
    rdClean = []

    for n in read_data:
        a = n.replace('\n', '')
        rdClean.append(a)
    print rdClean

    PNs = []
    InstanceIDs = []

    for itemName in rdClean:
        if 'ECS' in itemName:
            InstanceIDs.append(itemName)
        else:
            PNs.append(itemName)

    InstanceIDs = filter(len, InstanceIDs)
    PNs = filter(len, PNs)
    print PNs
    print InstanceIDs
    return (PNs, InstanceIDs)

def IRMer_nonc():

    global itemProd
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    state = True

    for prod in xrange(1, 5):
        selection1 = productDocument1.Selection
        selection1.Clear()
        #blocks_qty = 0
        product_inwork = products1.Item(prod)
        print product_inwork.PartNumber
        products_inwork = product_inwork.Products
  #UPPERS NON-CONSTANT
        for det in xrange(1, products_inwork.Count+1):
            product_inwork_nonc = products_inwork.Item(det)
            print product_inwork_nonc.PartNumber
            if 'FAIRING' in str(product_inwork_nonc.PartNumber):
                Instantiate_fairing(product_inwork_nonc)
            else:
                #blocks_qty += 1
                products_inwork_nonc = product_inwork_nonc.Products
                irm_type = 'UPR'
                if state:
                    product_forpaste_upr = add_new_irm(irm_type)
                for det_deep in xrange(1, 3):
                    product_highl_inwork_nonc = products_inwork_nonc.Item(det_deep)
                    selection1.Add(product_highl_inwork_nonc)

                #if blocks_qty != 0:
                paste(selection1, product_forpaste_upr)
                #selector(selection1, products1)
                itemProd += 1
                blocks_qty = 0
            
                #else:
                    #pass
  # LOWERS NON-CONSTANT:

                selection1 = productDocument1.Selection
                selection1.Clear()
                irm_type = 'LWR'
                if state:
                    product_forpaste_lwr = add_new_irm(irm_type)
                for det_deep in xrange(3, products_inwork_nonc.Count):
                    product_highl_inwork_nonc = products_inwork_nonc.Item(det_deep)
                    selection1.Add(product_highl_inwork_nonc)

                #if blocks_qty != 0:
                paste(selection1, product_forpaste_lwr)
                #selector(selection1, products1)
                itemProd += 1
                state = False
        #else:
            #pass

def IRMer_const():
    global constSize
    global itemProd
    global order_of_templete_product
    bin_breaks = [561, 690 + 120, 897 + 120, 1089 + 120, 1290 + 120, 1401 + 96 + 120, 1560 + 240]
    breaker = 0
    num = 0
    state = True
    initial_side = 'LH'
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products

    for prod in xrange(5, constSize+5):
        selection1 = productDocument1.Selection
        selection1.Clear()
        product_inwork = products1.Item(prod)
        print product_inwork.PartNumber
        print product_inwork.name
#UPPERS:
        if 'FAIRING' in str(product_inwork.PartNumber):
            Instantiate_fairing(product_inwork)
        else:
            # determinates switch to RH side
            if initial_side == 'LH':
                if 'RH' in str(product_inwork.Name):
                    initial_side = 'RH'
                    num = 0
                    itemProd += 1
                    state = True
            products_inwork = product_inwork.Products
            irm_type = 'UPR'
            if state:
                product_forpaste_upr = add_new_irm(irm_type)
            for det_deep in xrange(1, 3):
                product_highl_inwork = products_inwork.Item(det_deep)
                print product_highl_inwork.name
                selection1.Add(product_highl_inwork)

            paste(selection1, product_forpaste_upr)
            #selector(selection1, products1)

#LOWERS:
            selection1 = productDocument1.Selection
            selection1.Clear()
            irm_type = 'LWR'
            if state:
                product_forpaste_lwr = add_new_irm(irm_type)
            for det_deep in xrange(3, products_inwork.Count):
                product_highl_inwork = products_inwork.Item(det_deep)
                selection1.Add(product_highl_inwork)

            #selector(selection1, products1, True)
            paste(selection1, product_forpaste_lwr)
            order_of_templete_product += 1

            if bin_breaks[num] <= bin_breaker[breaker + 1]:
                if (num + 1) == len(bin_breaks):
                    num = 0
                else:
                    num += 1
                if (breaker + 2) < len(bin_breaker):
                    breaker += 1
                itemProd += 2
                state = True

            else:
                if (breaker + 2) < len(bin_breaker):
                    breaker += 1
                state = False
                continue

def add_new_irm(irm_type):
    global order_of_templete_product
    global itemProd
    global order_of_new_product
    global dash_number
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    selection1 = productDocument1.Selection
    selection1.Clear()
    pn = 'IR830ZXXXX-' + str(dash_number)
    id = 'ECS_' + irm_type + '_AIR_DIST_INSL_' + str(dash_number)
    product_forpaste = products1.AddNewComponent("Product", pn)
    product_forpaste.name = id
    order_of_new_product = dash_number - 995 + constSize
    dash_number += 1
    return product_forpaste


def paste(selection1, product_forpaste):

    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    selection1.Copy()
    selection1.Clear()
    selection2 = productDocument1.Selection
    selection2.Clear()
    selection2.add(product_forpaste)
    selection2.Paste()
    selection2.Clear()


def Instantiate_fairing(product_inwork_nonc):
    global order_of_templete_product
    global itemProd
    global order_of_new_product
    global dash_number
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    selection1 = productDocument1.Selection
    selection1.Clear()

    pn = 'IR830ZXXXX-' + str(dash_number)
    id = 'ECS_OMF_AIR_DIST_INSL_' + str(dash_number)
    product_forpaste = products1.AddNewComponent("Product", pn)
    product_forpaste.name = id
    order_of_new_product = dash_number - 995 + constSize
    dash_number += 1

    copy_from_name = product_inwork_nonc.Name
    print copy_from_name
    products_inwork_nonc = product_inwork_nonc.Products

    size = product_inwork_nonc.Name[:2]
    if 'ARCH' in product_inwork_nonc.Name[-7:]:

            arch = True
    else:

            arch = False

    if 'LH' in product_inwork_nonc.Name:
        side = 'LH'
    else:
        side = 'RH'

    for det_deep in xrange(1, 3):
            product_highl_inwork_nonc = products_inwork_nonc.Item(det_deep)
            selection1.Add(product_highl_inwork_nonc)

    selection1.Copy()
    selection1.Clear()
    selection2 = productDocument1.Selection
    selection2.Clear()
    selection2.add(product_forpaste)
    selection2.Paste()
    selection2.Clear()
    itemProd += 1
    order_of_templete_product += 1

    for_paste = products1.Item(order_of_new_product)
    product_name = for_paste.name
    product_pn = for_paste.PartNumber
    carm_pn = product_pn[2:]
    carm_name = product_name + '_CARM'
    print carm_pn
    print carm_name

    if make_carms == True:

        carm_instance = CARM(carm_pn, carm_name, side, order_of_new_product, order_of_templete_product)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        carm_instance.add_carm_as_external_component()
        carm_instance.change_inst_id()
        carm_instance.set_parameters(sta_value_pairs, size)
        carm_instance.modif_ref_annotation(size)
        carm_instance.modif_sta_annotation(sta_values_fake)
        carm_instance.copy_ref_surface_and_paste(size)
        #carm_instance.copy_bodies_and_paste('BACS12FA3K3')
        #carm_instance.copy_bodies_and_paste('FCM10F5CPS05WH')
        carm_instance.copy_jd1_fcm10f5cps05wh_and_paste(size)
        carm_instance.copy_jd2_bacs12fa3k3_and_paste(size, arch)
        carm_instance.create_jd_vectors(1)
        carm_instance.create_jd_vectors(2)
        carm_instance.copy_jd1_fcm10f5cps05wh_and_paste(size, 'vector')
        carm_instance.copy_jd2_bacs12fa3k3_and_paste(size, arch, 'vector')
        carm_instance.rename_vectors(1)
        carm_instance.rename_vectors(2)
        carm_instance.access_captures(4)
        carm_instance.add_jd_annotation('01', sta_value_pairs, size, side)
        carm_instance.access_captures(5)
        carm_instance.add_jd_annotation('02', sta_value_pairs, size, side)
        carm_instance.shift_camera(sta_value_pairs, size)
        carm_instance.access_captures(1)
        carm_instance.hide_unhide_captures('unhide', 1)
        carm_instance.activate_top_prod()

def part_number_collector():
    ICM_1.ApplyWorkMode(2)
    replacedDetails = []
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    print 'IRMs using the following part references:'

    for prod in xrange(1, 5):
        product_to_replace = products1.Item(prod)
        products_to_replace = product_to_replace.Products

        for det in xrange(1, products_to_replace.Count+1):
            product_act_to_replace_nonc = products_to_replace.Item(det)
            products_act_to_replace_nonc = product_act_to_replace_nonc.Products

            for det_deep in xrange(1, products_act_to_replace_nonc.Count):
                product_act_to_replace = products_act_to_replace_nonc.Item(det_deep)
                if "CA" in str(product_act_to_replace.PartNumber):
                    continue
                else:
                    replacedDetail = str(product_act_to_replace.PartNumber)
                    if replacedDetail in replacedDetails:
                        continue
                    else:
                        replacedDetails.append(replacedDetail)
                        print replacedDetail

    for prod in xrange(5, products1.Count+1):
        product_to_replace = products1.Item(prod)
        products_to_replace = product_to_replace.Products

        for det in xrange(1, products_to_replace.Count):
            product_act_to_replace = products_to_replace.Item(det)
            if "CA" in str(product_act_to_replace.PartNumber):
                continue
            else:
                replacedDetail = str(product_act_to_replace.PartNumber)
            if replacedDetail in replacedDetails:
                continue
            else:
                replacedDetails.append(replacedDetail)
                print replacedDetail

def Deleter():
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    selection1 = productDocument1.Selection
    selection1.Clear()

    for prod in xrange(1, constSize+5):
        product_inwork = products1.Item(prod)
        print product_inwork.PartNumber + 'deleted'
        print product_inwork.name + 'deleted'
        selection1.Add(product_inwork)

    selection1.Delete()
    selection1.Clear()

def nonconst_comps_inst():
    ICM = CATIA.ActiveDocument
    oFileSys = CATIA.FileSystem
    ICM_1 = ICM.Product
    ICM_Products = ICM_1.Products

    global new_component1
    new_component1 = ICM_Products.AddNewComponent("Product", 'non-constant_41_LH')
    global ICM_Sec41_LH_Products
    ICM_Sec41_LH_Products = new_component1.Products

    global new_component2
    new_component2 = ICM_Products.AddNewComponent("Product", 'non-constant_41_RH')
    global ICM_Sec41_RH_Products
    ICM_Sec41_RH_Products = new_component2.Products

    global new_component3
    new_component3 = ICM_Products.AddNewComponent("Product", 'non-constant_47_LH')
    global ICM_Sec47_LH_Products
    ICM_Sec47_LH_Products = new_component3.Products

    global new_component4
    new_component4 = ICM_Products.AddNewComponent("Product", 'non-constant_47_RH')
    global ICM_Sec47_RH_Products
    ICM_Sec47_RH_Products = new_component4.Products

def AddComponent(s, side, section, location, plug_value):
    #path = '\\\\FIL-MOW01-01\\787Payloads\\IRC\\SYSTEM_Int\\APIE_TOOL\\APIE_Tool2\\LIBRARY_NOGEOM_NOICM'

    extention = '.CATProduct'
    global sta_value_pairs
    global sta_values_fake
    x_coord = Inch_to_mm(465)
    x_coord_nonconstant = Inch_to_mm(0)
    fake_coord_nonconstant_41 = Inch_to_mm(459)
    if plug_value == 240:
        fake_coord_nonconstant_47 = Inch_to_mm(1863)
    elif plug_value == 456:
        fake_coord_nonconstant_47 = Inch_to_mm(2079)
    elif plug_value == 0:
        fake_coord_nonconstant_47 = Inch_to_mm(1623)

    if plug_value == 0:
        door2_coord = 0
    elif plug_value == 456:
        door2_coord = 240
    else:
        door2_coord = 120

    if side == 'LH' and location == 'middle':
        iteration = 0
    elif side == 'RH' and location == 'middle':
        iteration = 100
    elif side == 'LH' and location == 'nose':
        iteration = 200
    elif side == 'RH' and location == 'nose':
        iteration = 300
    elif side == 'LH' and location == 'tail':
        iteration = 400
    elif side == 'RH' and location == 'tail':
        iteration = 500

    if location == 'tail':
        angle = 3.125
    else:
        angle = 5

    rad = math.radians(angle)
    print angle

    index = 0

    for number in s:

        nozzl_type = 'ECO'
        dow_type = 'DWNR_STD-STRT'
        ligval_ammount = 1
        Arch = ''

        bins = ['36', '42', '48', '362', '422', '482']
        bin_twenty_four = ['24', '242', '2432', '243']

        if number in bins:
            stowbin = True
            btype = 'BIN'
        elif number in bin_twenty_four:
            stowbin = 'twenty_four'
            dow_type = 'DWNR_JOG-STRT'
            btype = 'BIN'
        else:
            stowbin = False
            btype = 'FAIRING'

        if str(number) == 'door':
            x_coord = Inch_to_mm(693 + door2_coord)
            index += 1
            continue

        else:

            Rotate5 = [0.996194698, -0.087155742, 0, 0.087155742, 0.996194698, 0, 0, 0, 1, Inch_to_mm(466.61647022),
                       Inch_to_mm(0.08471639), 0]
            Rotate185 = [-0.996194698, -0.087155742, 0, 0.087155742, -0.996194698, 0, 0, 0, 1, Inch_to_mm(466.61647018),
                         Inch_to_mm(-0.084716377), 0]
            Rotate_5 = [0.998512978, 0.054514501, 0, -0.054514501, 0.998512978, 0, 0, 0, 1,
                        Inch_to_mm(1618.61663822 + plug_value), Inch_to_mm(0.17865996), 0]
            Rotate_185 = [-0.998512978, 0.054514501, 0, -0.054514501, -0.998512978, 0, 0, 0, 1,
                          Inch_to_mm(1618.61663822 + plug_value), Inch_to_mm(-0.17865996), 0]

            print int(number)  # check

            # checking area around DOOR 2:

            if index != (len(s) - 1) and (s[index + 1] == 'door' or s[index - 1] == 'door'):

                if (side == 'LH' and s[index + 1] == 'door') or (side == 'RH' and s[index - 1] == 'door'):

                    print 'RH door2'

                    Arch = 'ARCH'

                    if int(number) == 24:
                        dow_type = 'DWNR_JOG-RGHT'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Twenty_four_arch_RH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 243:
                        number = '30'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Twenty_four_arch_EXT_RH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 36:
                        dow_type = 'DWNR_JOG-STRT'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_six_arch_RH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 42:

                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fourty_two_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 18:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Eighteen_arch_RH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 30:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_arch_RH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 54:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fifty_four_arch_RH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 60:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Sixty_arch_RH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 241:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Twenty_four_fairing_arch_RH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '24'

                    elif int(number) == 361:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_six_fairing_arch_RH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '36'

                    elif int(number) == 421:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fourty_two_fairing_arch_RH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '42'

                    elif int(number) == 481:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fourty_eight_fairing_arch_RH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '48'

                    #PREMIUM:

                    elif int(number) == 242:
                        dow_type = 'DWNR_JOG-RGHT'
                        number = '24'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Twenty_four_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 2432:
                        number = '30'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Twenty_four_arch_EXT_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 182:
                        number = '18'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Eighteen_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 302:
                        number = '30'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 542:
                        number = '54'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fifty_four_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 602:
                        number = '60'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Sixty_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 2412:
                        number = '24'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Twenty_four_fairing_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '24'

                    elif int(number) == 3612:
                        number = '36'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_six_fairing_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '36'

                    elif int(number) == 4212:
                        number = '42'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fourty_two_fairing_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '42'

                    elif int(number) == 4812:
                        number = '48'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fourty_eight_fairing_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '48'

                    else:
                        x_coord += Inch_to_mm(int(number))
                        iteration += 1
                        index += 1
                        continue

                elif (side == 'LH' and s[index - 1] == 'door') or (side == 'RH' and s[index + 1] == 'door'):

                    print 'LH door2'

                    Arch = 'ARCH'

                    if int(number) == 24:
                        dow_type = 'DWNR_JOG-LEFT'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Twenty_four_arch_LH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 243:
                        number = '30'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Twenty_four_arch_EXT_LH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 36:
                        dow_type = 'DWNR_JOG-STRT'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_six_arch_LH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 42:

                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fourty_two_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 18:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Eighteen_arch_LH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 30:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_arch_LH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 54:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fifty_four_arch_LH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 60:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Sixty_arch_LH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 241:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Twenty_four_fairing_arch_LH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '24'

                    elif int(number) == 361:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_six_fairing_arch_LH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '36'

                    elif int(number) == 421:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fourty_two_fairing_arch_LH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '42'

                    elif int(number) == 481:
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fourty_eight_fairing_arch_LH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '48'

                        #PREM:

                    elif int(number) == 242:
                        dow_type = 'DWNR_JOG-LEFT'
                        number = '24'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Twenty_four_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 2432:
                        number = '30'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Twenty_four_arch_EXT_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 182:
                        number = '18'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Eighteen_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 302:
                        number = '30'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 542:
                        number = '54'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fifty_four_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 602:
                        number = '60'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Sixty_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 2412:
                        number = '24'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Twenty_four_fairing_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '24'

                    elif int(number) == 3612:
                        number = '36'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_six_fairing_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '36'

                    elif int(number) == 4212:
                        number = '42'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fourty_two_fairing_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '42'

                    elif int(number) == 4812:
                        number = '48'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fourty_eight_fairing_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '48'

                    else:
                        x_coord += Inch_to_mm(int(number))
                        iteration += 1
                        index += 1
                        continue

                        #  NOT around DOOR 2:

            elif (location == 'nose' and stowbin is not True and stowbin != 'twenty_four') or location == 'middle':

                if int(number) == 24:

                    iteration += 1
                    index += 1
                    if index != len(s) and side == 'RH' and s[index] == '72' or index != len(s) and side == 'LH' and s[
                                index - 2] == '72':
                        PartDocPath = path + '\Twenty_four_DR3_LH_solids'
                        dow_type = 'DWNR_JOG-LEFT'
                    elif index != len(s) and side == 'RH' and s[index - 2] == '72' or index != len(
                            s) and side == 'LH' and s[index] == '72':
                        PartDocPath = path + '\Twenty_four_DR3_RH_solids'
                        dow_type = 'DWNR_JOG-RGHT'
                    else:
                        PartDocPath = path + '\Twenty_four_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 243:
                    number = '30'
                    iteration += 1
                    index += 1
                    if index != len(s) and side == 'RH' and s[index] == '72' or index != len(s) and side == 'LH' and s[
                                index - 2] == '72':
                        PartDocPath = path + '\Twenty_four_EXT_DR3_LH_solids'
                        dow_type = 'DWNR_JOG-STRT'
                        ligval_ammount = 2
                    elif index != len(s) and side == 'RH' and s[index - 2] == '72' or index != len(
                            s) and side == 'LH' and s[index] == '72':
                        PartDocPath = path + '\Twenty_four_EXT_DR3_RH_solids'
                        dow_type = 'DWNR_JOG-STRT'
                        ligval_ammount = 2
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 36:

                    iteration += 1
                    index += 1
                    if index != len(s) and side == 'RH' and s[index] == '72' or index != len(s) and side == 'LH' and s[
                                index - 2] == '72':
                        PartDocPath = path + '\Thirty_six_DR3_solids'
                        dow_type = 'DWNR_JOG-STRT'
                    elif index != len(s) and side == 'RH' and s[index - 2] == '72' or index != len(
                            s) and side == 'LH' and s[index] == '72':
                        PartDocPath = path + '\Thirty_six_DR3_solids'
                        dow_type = 'DWNR_JOG-STRT'
                    else:
                        PartDocPath = path + '\Thirty_six_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 42:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_two_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 48:

                    iteration += 1
                    index += 1
                    if STAvalue(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'LH' or STAvalue(
                            x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'LH':
                        PartDocPath = path + '\Fourty_eight_horseshoe_solids_LH'
                    elif STAvalue(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'RH' or STAvalue(
                            x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'RH':
                        PartDocPath = path + '\Fourty_eight_horseshoe_solids_RH'
                    else:
                        PartDocPath = path + '\Fourty_eight_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 12:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Twelve_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 18:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Eighteen_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 30:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Thirty_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 54:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fifty_four_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 60:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Sixty_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 72:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Seventy_two_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 241:

                    number = '24'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Twenty_four_fairing_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 361:

                    number = '36'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Thirty_six_fairing_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 421:

                    number = '42'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_two_fairing_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 481:

                    number = '48'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_eight_fairing_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                    #Premium plenums:

                elif int(number) == 242:

                    number = '24'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    if index != len(s) and side == 'RH' and s[index] == '72' or index != len(s) and side == 'LH' and s[
                                index - 2] == '72':
                        PartDocPath = path + '\Twenty_four_DR3_LH_solids_pr'
                        dow_type = 'DWNR_JOG-LEFT'
                    elif index != len(s) and side == 'RH' and s[index - 2] == '72' or index != len(
                            s) and side == 'LH' and s[index] == '72':
                        PartDocPath = path + '\Twenty_four_DR3_RH_solids_pr'
                        dow_type = 'DWNR_JOG-RGHT'
                    else:
                        PartDocPath = path + '\Twenty_four_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 2432:
                    number = '30'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    if index != (len(s) - 1) and side == 'RH' and s[index] == '72' or index != (
                        len(s) - 1) and side == 'LH' and s[index - 2] == '72':
                        PartDocPath = path + '\Twenty_four_EXT_DR3_LH_solids_pr'
                        dow_type = 'DWNR_JOG-STRT'
                        ligval_ammount = 2
                    elif index != (len(s) - 1) and side == 'RH' and s[index - 2] == '72' or index != (
                        len(s) - 1) and side == 'LH' and s[index] == '72':
                        PartDocPath = path + '\Twenty_four_EXT_DR3_RH_solids_pr'
                        dow_type = 'DWNR_JOG-STRT'
                        ligval_ammount = 2
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 362:

                    number = '36'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Thirty_six_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 422:

                    number = '42'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_two_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 482:

                    number = '48'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    if STAvalue(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'LH' or STAvalue(
                            x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'LH':
                        PartDocPath = path + '\Fourty_eight_horseshoe_solids_LH'
                    elif STAvalue(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'RH' or STAvalue(
                            x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'RH':
                        PartDocPath = path + '\Fourty_eight_horseshoe_solids_RH'
                    else:
                        PartDocPath = path + '\Fourty_eight_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 122:

                    number = '12'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Twelve_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 182:

                    number = '18'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Eighteen_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 302:

                    number = '30'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Thirty_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 542:

                    number = '54'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fifty_four_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 602:

                    number = '60'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Sixty_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 722:

                    number = '72'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Seventy_two_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 2412:

                    number = '24'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Twenty_four_fairing_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 3612:

                    number = '36'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Thirty_six_fairing_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 4212:

                    number = '42'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_two_fairing_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 4812:

                    number = '48'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_eight_fairing_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                else:
                    x_coord += Inch_to_mm(int(number))
                    iteration += 1
                    index += 1
                    continue

                    #SECTION 41:

            elif location == 'nose' and stowbin is True or stowbin == 'twenty_four' and location == 'nose':

                if int(number) == 24:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Twenty_four_solids_sec41'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 36:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Thirty_six_solids_sec41'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 42:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_two_solids_sec41'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 48:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_eight_solids_sec41'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 242:

                    number = '24'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Twenty_four_solids_sec41_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)


                elif int(number) == 362:

                    number = '36'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Thirty_six_solids_sec41_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 422:

                    number = '42'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_two_solids_sec41_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 482:

                    number = '48'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_eight_solids_sec41_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                    #SECTION 47:

            else:

                if int(number) == 24:
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Twenty_four_solids_sec47'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 36:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Thirty_six_solids_sec47'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 42:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_two_solids_sec47'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 48:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_eight_solids_sec47'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 12:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Twelve_solids_sec47'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 18:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Eighteen_solids_sec47'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 30:

                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Thirty_solids_sec47'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 241:

                    number = '24'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Twenty_four_fairing_solids_sec47'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 361:

                    number = '36'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Thirty_six_fairing_solids_sec47'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 421:

                    number = '42'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_two_fairing_solids_sec47'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 481:

                    number = '48'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_eight_fairing_solids_sec47'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                else:
                    x_coord += Inch_to_mm(int(number))
                    iteration += 1
                    index += 1
                    continue

                    #for lower plenums:

            if stowbin is True or 'twenty_four':
                if number == '36':
                    L_PL_size1 = '36'
                    L_PL_size2 = '36'
                elif number == '42':
                    L_PL_size1 = '42'
                    L_PL_size2 = '42'
                elif number == '48':
                    L_PL_size1 = '48'
                    L_PL_size2 = '48'
                else:
                    L_PL_size1 = '24'

            if section == 'constant':

                NewComponent = ICM_Products.AddExternalComponent(PartDoc)
                PartDoc.Close()
                oFileSys.DeleteFile(PartDocPath1)
                NewComponent.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                NewComponent.Name = str(number) + 'IN STA ' + STAvalue(x_coord, plug_value) + ' ' + side + Arch
                trouble2 = mm_to_Inch(x_coord)
                bin_breaker.append(int(trouble2))
                sta_values_fake.append(STAvalue(x_coord, plug_value))
                sta_value_pairs.append(x_coord)
                print sta_value_pairs
                print sta_values_fake

                print bin_breaker

                RenamingTool = NewComponent.ReferenceProduct
                PlenumAssy = RenamingTool.Products.Item(1)
                LING_VAL = RenamingTool.Products.Item(2)
                if ligval_ammount == 2:
                    LING_VAL2 = RenamingTool.Products.Item(3)

                if stowbin is True:
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Downer2 = RenamingTool.Products.Item(5)
                    Lower_Plenum1.name = 'CONST_' + L_PL_size1 + 'LWPLEN_STA' + STAvalue(x_coord, plug_value) + '_' + \
                                         side[0]
                    Lower_Downer1.name = dow_type + '_STA' + STAvalue(x_coord, plug_value) + '_' + side[0] + '1'
                    Lower_Downer2.name = dow_type + '_STA' + STAvalue(x_coord, plug_value) + '_' + side[0] + '2'



                elif stowbin == 'twenty_four':
                    if ligval_ammount == 2:
                        Lower_Plenum1 = RenamingTool.Products.Item(4)
                        Lower_Downer1 = RenamingTool.Products.Item(5)
                        Lower_Plenum1.name = 'CONST_' + L_PL_size1 + 'LWPLEN_STA' + STAvalue(x_coord,
                                                                                             plug_value) + '_' + side[0]
                        Lower_Downer1.name = dow_type + '_STA' + STAvalue(x_coord, plug_value) + '_' + side[0]



                    elif ligval_ammount == 1:
                        Lower_Plenum1 = RenamingTool.Products.Item(3)
                        Lower_Downer1 = RenamingTool.Products.Item(4)
                        Lower_Plenum1.name = 'CONST_' + L_PL_size1 + 'LWPLEN_STA' + STAvalue(x_coord,
                                                                                             plug_value) + '_' + side[0]
                        Lower_Downer1.name = dow_type + '_STA' + STAvalue(x_coord, plug_value) + '_' + side[0]

                PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + STAvalue(x_coord, plug_value) + '_' + \
                                  side[0]
                if len(PlenumAssy.name) > 24:
                    PlenumAssy.name = str(number) + nozzl_type + 'ASSY_' + 'STA' + STAvalue(x_coord, plug_value) + '_' + \
                                      side[0]
                if ligval_ammount == 1:
                    LING_VAL.name = 'OB_BIN_LIGVAL_STA' + STAvalue(x_coord, plug_value) + '_' + side[0]
                    if len(LING_VAL.name) > 24:
                        LING_VAL.name = 'OB_LIGVAL_STA' + STAvalue(x_coord, plug_value) + '_' + side[0]
                elif ligval_ammount == 2:
                    LING_VAL.name = 'OB_BIN_LIGVAL_STA' + STAvalue(x_coord, plug_value) + '_' + side[0] + '1'
                    if len(LING_VAL.name) > 24:
                        LING_VAL.name = 'OB_LIGVAL_STA' + STAvalue(x_coord, plug_value) + '_' + side[0] + '1'
                    LING_VAL2.name = 'OB_BIN_LIGVAL_STA' + STAvalue(x_coord, plug_value) + '_' + side[0] + '2'
                    if len(LING_VAL2.name) > 24:
                        LING_VAL.name = 'OB_LIGVAL_STA' + STAvalue(x_coord, plug_value) + '_' + side[0] + '2'


            elif section == 'nonconstant' and side == 'LH' and location == 'nose':

                NewComponent = ICM_Sec41_LH_Products.AddExternalComponent(PartDoc)
                PartDoc.Close()
                oFileSys.DeleteFile(PartDocPath1)
                RenamingToolProd = new_component1.ReferenceProduct
                Prod = RenamingToolProd.Products.Item(index)
                Prod.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                Prod.Name = str(number) + 'IN STA ' + STAvalue(
                    (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))),
                    plug_value) + ' ' + side
                RenamingTool = NewComponent.ReferenceProduct
                PlenumAssy = RenamingTool.Products.Item(1)
                LING_VAL = RenamingTool.Products.Item(2)

                if stowbin is True:
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Downer2 = RenamingTool.Products.Item(5)
                    Lower_Plenum1.name = 'SEC41_' + L_PL_size1 + 'LWPLEN_STA' + STAvalue(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0]
                    Lower_Downer1.name = dow_type + '_STA' + STAvalue(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0] + '1'
                    Lower_Downer2.name = dow_type + '_STA' + STAvalue(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0] + '2'

                elif stowbin == 'twenty_four':
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Plenum1.name = 'SEC41_' + L_PL_size1 + 'LWPLEN_STA' + STAvalue(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0]
                    Lower_Downer1.name = dow_type + '_STA' + STAvalue(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0]

                PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + STAvalue(
                    (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + \
                                  side[0]
                LING_VAL.name = 'OB_BIN_LIGVAL_STA' + STAvalue(
                    (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + \
                                side[0]
                NewComponent.Move.Apply(Rotate5)

            elif section == 'nonconstant' and side == 'RH' and location == 'nose':

                NewComponent = ICM_Sec41_RH_Products.AddExternalComponent(PartDoc)
                PartDoc.Close()
                oFileSys.DeleteFile(PartDocPath1)
                RenamingToolProd = new_component2.ReferenceProduct
                Prod = RenamingToolProd.Products.Item(index)
                Prod.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                Prod.Name = str(number) + 'IN STA ' + STAvalue(
                    (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))),
                    plug_value) + ' ' + side
                RenamingTool = NewComponent.ReferenceProduct
                PlenumAssy = RenamingTool.Products.Item(1)
                LING_VAL = RenamingTool.Products.Item(2)

                if stowbin is True:
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Downer2 = RenamingTool.Products.Item(5)
                    Lower_Plenum1.name = 'SEC41_' + L_PL_size1 + 'LWPLEN_STA' + STAvalue(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0]
                    Lower_Downer1.name = dow_type + '_STA' + STAvalue(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0] + '1'
                    Lower_Downer2.name = dow_type + '_STA' + STAvalue(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0] + '2'

                elif stowbin == 'twenty_four':
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Plenum1.name = 'SEC41_' + L_PL_size1 + 'LWPLEN_STA' + STAvalue(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0]
                    Lower_Downer1.name = dow_type + '_STA' + STAvalue(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0]

                PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + STAvalue(
                    (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + \
                                  side[0]
                LING_VAL.name = 'OB_BIN_LIGVAL_STA' + STAvalue(
                    (fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + \
                                side[0]
                NewComponent.Move.Apply(Rotate185)

            elif section == 'nonconstant' and side == 'LH' and location == 'tail':

                NewComponent = ICM_Sec47_LH_Products.AddExternalComponent(PartDoc)
                PartDoc.Close()
                oFileSys.DeleteFile(PartDocPath1)
                RenamingToolProd = new_component3.ReferenceProduct
                Prod = RenamingToolProd.Products.Item(index)
                Prod.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                Prod.Name = str(number) + 'IN STA ' + STAvalue((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                               plug_value) + ' ' + side
                RenamingTool = NewComponent.ReferenceProduct
                PlenumAssy = RenamingTool.Products.Item(1)
                Felt = RenamingTool.Products.Item(2)

                if stowbin is True:
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Downer2 = RenamingTool.Products.Item(5)
                    Lower_Plenum1.name = 'CONST_' + L_PL_size1 + 'LWPLEN_STA' + STAvalue(
                        (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                    Lower_Downer1.name = dow_type + '_STA' + STAvalue((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                      plug_value) + '_' + side[0] + '1'
                    Lower_Downer2.name = dow_type + '_STA' + STAvalue((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                      plug_value) + '_' + side[0] + '2'



                elif stowbin == 'twenty_four':
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Plenum1.name = 'CONST_' + L_PL_size1 + 'LWPLEN_STA' + STAvalue(
                        (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                    Lower_Downer1.name = dow_type + '_STA' + STAvalue((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                      plug_value) + '_' + side[0]

                PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + STAvalue(
                    (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                Felt.name = 'UPR_FELT_' + str(number) + 'IN_STA' + STAvalue(
                    (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                NewComponent.Move.Apply(Rotate_5)

            elif section == 'nonconstant' and side == 'RH' and location == 'tail':

                NewComponent = ICM_Sec47_RH_Products.AddExternalComponent(PartDoc)
                PartDoc.Close()
                oFileSys.DeleteFile(PartDocPath1)
                RenamingToolProd = new_component4.ReferenceProduct
                Prod = RenamingToolProd.Products.Item(index)
                Prod.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                Prod.Name = str(number) + 'IN STA ' + STAvalue((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                               plug_value) + ' ' + side
                RenamingTool = NewComponent.ReferenceProduct
                PlenumAssy = RenamingTool.Products.Item(1)
                Felt = RenamingTool.Products.Item(2)

                if stowbin is True:
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Downer2 = RenamingTool.Products.Item(5)
                    Lower_Plenum1.name = 'CONST_' + L_PL_size1 + 'LWPLEN_STA' + STAvalue(
                        (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                    Lower_Downer1.name = dow_type + '_STA' + STAvalue((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                      plug_value) + '_' + side[0] + '1'
                    Lower_Downer2.name = dow_type + '_STA' + STAvalue((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                      plug_value) + '_' + side[0] + '2'



                elif stowbin == 'twenty_four':
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Plenum1.name = 'CONST_' + L_PL_size1 + 'LWPLEN_STA' + STAvalue(
                        (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                    Lower_Downer1.name = dow_type + '_STA' + STAvalue((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                      plug_value) + '_' + side[0]

                PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + STAvalue(
                    (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                Felt.name = 'UPR_FELT_' + str(number) + 'IN_STA' + STAvalue(
                    (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                NewComponent.Move.Apply(Rotate_185)

            if location == 'nose':
                x_coord_nonconstant -= Inch_to_mm(int(number))

            x = x_coord_nonconstant * math.cos(rad)
            y = x_coord_nonconstant * math.sin(rad)

            position = [1, 0, 0, 0, 1, 0, 0, 0, 1, x_coord, 0, 0]
            position_non = [1, 0, 0, 0, 1, 0, 0, 0, 1, x, -y, 0]
            position_non_RH = [1, 0, 0, 0, 1, 0, 0, 0, 1, x + (Inch_to_mm(int(number)) * math.cos(rad)),
                               y + (Inch_to_mm(int(number)) * math.sin(rad)), 0]
            position90 = [-1, 0, 0, 0, -1, 0, 0, 0, 1, x_coord + Inch_to_mm(int(number)), 0, 0]  # 90 deg rotation
            position_non_47 = [1, 0, 0, 0, 1, 0, 0, 0, 1, x, y, 0]
            position_non_47_RH = [1, 0, 0, 0, 1, 0, 0, 0, 1, x + (Inch_to_mm(int(number)) * math.cos(rad)),
                                  (y + (Inch_to_mm(int(number)) * math.sin(rad))) * (-1), 0]

            if side == 'LH' and section == 'constant':
                #NewComponentRef = NewComponent.ReferenceProduct
                NewComponent.Move.Apply(position)
                print side
                print x_coord
            elif side == 'RH' and section == 'constant':
                NewComponent.Move.Apply(position90)
                print side
                print x_coord
            elif section == 'nonconstant' and side == 'LH' and location == 'nose':
                NewComponent.Move.Apply(position_non)
                print section
                print x_coord_nonconstant
            elif section == 'nonconstant' and side == 'RH' and location == 'nose':
                NewComponent.Move.Apply(position_non_RH)
                print section
                print x_coord_nonconstant
            elif section == 'nonconstant' and side == 'LH' and location == 'tail':
                NewComponent.Move.Apply(position_non_47)
                x_coord_nonconstant += Inch_to_mm(int(number))
                print section
                print x, y
                print x_coord_nonconstant
            elif section == 'nonconstant' and side == 'RH' and location == 'tail':
                NewComponent.Move.Apply(position_non_47_RH)
                x_coord_nonconstant += Inch_to_mm(int(number))
                print section
                print x, y
                print x_coord_nonconstant

            x_coord += Inch_to_mm(int(number))


if __name__ == '__main__':
    root = Tkinter.Tk()
    TkFileDialogExample(root).pack()
    root.mainloop()
