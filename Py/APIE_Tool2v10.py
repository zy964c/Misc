import Tkinter
import Tkconstants
import tkFileDialog
import tkMessageBox
import math
import time
import win32com.client
from carm3 import*

start = time.time()
CATIA = win32com.client.Dispatch('catia.application')

try:
    ICM = CATIA.ActiveDocument
except:
    ICM = CATIA.Documents.Add('Product')

oFileSys = CATIA.FileSystem

ICM_1 = ICM.Product
ICM_Products = ICM_1.Products

global angle
angle = 0

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

global make_carms_UI
make_carms_UI = False

global dash_number
dash_number = 1000

global enovia_connected
enovia_connected = True

global create_irms
create_irms = True

global plug
plug = 240

global parameters_set
parameters_set = False

global path
# path = '\\\\nw\\data\\irc-kmapi\\IRC_KBE\\Engineering_Automation\\ECS_Tool\\LIBRARY_NOGEOM_ICM2'
# path = 'C:\\Temp\\zy964c\\11.03.2015_version\\LIBRARY_NOGEOM_ICM2'
path = 'E:\\ECS-tool\\LIBRARY_NOGEOM_ICM2'


var_import = [path]


class TkFileDialogExample(Tkinter.Frame):

    """
    Implements GUI
    """

    def __init__(self, root):

        Tkinter.Frame.__init__(self, root)

        self.plug = Tkinter.IntVar()
        root.geometry("740x380")
        self.repl = Tkinter.IntVar()
        self.enovia = Tkinter.IntVar()
        self.irm = Tkinter.IntVar()

        # options for buttons
        button_opt = {'fill': Tkconstants.BOTH, 'padx': 5, 'pady': 5}

        # define buttons
        self.sp = Tkinter.StringVar()
        self.lp = Tkinter.StringVar()

        Tkinter.Button(self, text='Choose a bin run', command=self.askopenfile).pack(**button_opt)
        Tkinter.Radiobutton(root, text="787-8", variable=self.plug, value=0).pack(**button_opt)
        Tkinter.Radiobutton(root, text="787-9", variable=self.plug, value=240).pack(**button_opt)
        Tkinter.Radiobutton(root, text="787-10", variable=self.plug, value=456).pack(**button_opt)
        Tkinter.Label(root, text="Enter library path:").pack()
        Tkinter.Entry(root, textvariable=self.lp, width=100).pack()
        Tkinter.Label(root, text="Make CARMs").pack()
        Tkinter.Checkbutton(root, variable=self.repl).pack()
        Tkinter.Label(root, text="Do not replace library part references with ENOVIA part references").pack()
        Tkinter.Checkbutton(root, variable=self.enovia).pack()
        Tkinter.Label(root, text="Do not combine parts into IRMs").pack()
        Tkinter.Checkbutton(root, variable=self.irm).pack()
        Tkinter.Label(root, text="Please have an empty product in CATIA open before running the script").pack()

        self.plug.set(240)
        self.lp.set(path)

        # define options for opening or saving a file
        self.file_opt = options = {}
        options['defaultextension'] = '.txt'
        options['filetypes'] = [('all files', '.*'), ('text files', '.txt')]
        options['initialdir'] = 'C:\\'
        options['initialfile'] = 'myfile.txt'
        options['parent'] = root
        options['title'] = 'This is a title'

    def askopenfile(self):

        """
        Returns an opened file in read mode
        """

        global make_carms
        global make_carms_UI
        global enovia_connected
        global create_irms
        global plug

        f = tkFileDialog.askopenfile(parent=root, mode='rb', title='Choose a file')

        s = f.readlines()
        print s
        s_all = []
        state1 = True
        for element in s:
            if 'CTR' in element:
                state1 = False
                continue
            elif '#' in element and 'CTR' not in element:
                state1 = True
                continue
            elif '#' not in element and state1 is True:
                s_all.append(element.replace(' ', '').replace('fairing', '1').replace('premium', '2').replace('prem', '2').replace('EXT', '3').replace('\r\n', '').split(","))
            else:
                continue

        print s_all

        s1 = s_all[0]
        s2 = s_all[1]
        s3 = s_all[2]
        s4 = s_all[3]
        s5 = s_all[4]
        s6 = s_all[5]

        s1 = s1[::-1]
        s2 = s2[::-1]

        print s1
        print s2
        print s3
        print s4
        print s5
        print s6

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

        instantiate_nonconstant_components()

        if s3 != ['']:
            add_component(s3, 'LH', 'constant', 'middle', plug)
        if s4 != ['']:
            add_component(s4, 'RH', 'constant', 'middle', plug)
        if s1 != ['']:
            add_component(s1, 'LH', 'nonconstant', 'nose', 0)
        if s2 != ['']:
            add_component(s2, 'RH', 'nonconstant', 'nose', 0)
        if s5 != ['']:
            add_component(s5, 'LH', 'nonconstant', 'tail', plug)
        if s6 != ['']:
            add_component(s6, 'RH', 'nonconstant', 'tail', plug)

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
            make_carms = True
            make_carms_UI = True

        if self.enovia.get() == 1:
            enovia_connected = False

        if self.irm.get() == 1:
            create_irms = False

        if create_irms:
            if enovia_connected:
                part_number_collector()
                raw_input('Please bring to CATIA listed part references and press ENTER')
                replacer()
            reframe_all()
            irm_builder_constant()
            irm_builder_nonconstant()
            deleter()

        end = time.time()
        print end - start
        tkMessageBox.showinfo(title="APIE tool", message="Done")
        root.withdraw()
        root.destroy()

    def seed_carm_path(self):
        path_carm = str(self.lp.get())
        return path_carm


def inch_to_mm(distance):

    """
    :param distance: distance in inches
    :return: distance in mm
    """

    return distance * 25.4


def mm_to_inch(distance):

    """
    :param distance: distance in mm
    :return: distance in inches
    """

    return distance / 25.4


def sta_value(coord, plug_value):

    """
    :param coord: x coordinate in airplane coordinate system in mm
    :param plug_value: 240 for -9, 456 for -10, 0 for -8
    :return: str with format 'STA....' or 'STA....+...'
    """

    STA = '0'
    if plug_value == 240:
        if round(coord, 1) <= round(inch_to_mm(609), 1):
            STA = '0' + str(int(round(coord / 25.4)))
        elif round(coord, 1) > round(inch_to_mm(609), 1) and coord <= round(inch_to_mm(609 + 120), 1):
            STA = '0609+' + str(int(round(coord / 25.4 - 609)))
        elif round(coord, 1) > round(inch_to_mm(609 + 120), 1) and coord <= round(inch_to_mm(1401 + 120), 1):
            if (coord / 25.4 - 120) < 1000:
                STA = '0' + str(int(round(coord / 25.4 - 120)))
            else:
                STA = str(int(round(coord / 25.4 - 120)))
        elif round(coord, 1) > round(inch_to_mm(1401 + 120), 1) and coord <= round(inch_to_mm((1401 + 120) + 120), 1):
            STA = '1401+' + str(int(round(coord / 25.4 - (1401 + 120))))
        elif round(coord, 1) > round(inch_to_mm(1401 + 240), 1):
            STA = str(int(round(coord / 25.4 - 240)))

    elif plug_value == 456:
        if round(coord, 1) <= round(inch_to_mm(609), 1):
            STA = '0' + str(int(round(coord / 25.4)))
        elif round(coord, 1) > round(inch_to_mm(609), 1) and coord <= round(inch_to_mm(609 + 240), 1):
            STA = '0609+' + str(int(round(coord / 25.4 - 609)))
        elif round(coord, 1) > round(inch_to_mm(609 + 240), 1) and coord <= round(inch_to_mm(1401 + 240), 1):
            if (coord / 25.4 - 240) < 1000:
                STA = '0' + str(int(round(coord / 25.4 - 240)))
            else:
                STA = str(int(round(coord / 25.4 - 240)))
        elif round(coord, 1) > round(inch_to_mm(1401 + 240), 1) and coord <= round(inch_to_mm((1401 + 240) + 120), 1):
            STA = '1401+' + str(int(round(coord / 25.4 - (1401 + 240))))
        elif round(coord, 1) > round(inch_to_mm(1401 + 360), 1) and coord <= round(inch_to_mm(1618 + 360), 1):
            STA = str(int(round(coord / 25.4 - 360)))
        elif round(coord, 1) > round(inch_to_mm(1618 + 360), 1) and coord <= round(inch_to_mm((1618 + 360) + 96), 1):
            STA = '1618+' + str(int(round(coord / 25.4 - (1618 + 360))))
        elif round(coord, 1) > round(inch_to_mm(1618 + 360 + 96), 1):
            STA = str(int(round(coord / 25.4 - (360 + 96))))

    elif plug_value == 0:
        if int(round(coord / 25.4)) < 1000:
            STA = '0' + str(int(round(coord / 25.4)))
        else:
            STA = str(int(round(coord / 25.4)))
    return STA


def replacer():

    """
    Replaces part references from the library with equivalent ENOVIA references. revA, revB, revC lists should be
    maintained
    :return: nothing
    """

    ICM_1.ApplyWorkMode(2)
    replacedDetails = []
    revC = ["832Z4501-1"]
    revB = ["832Z4501-10", "832Z4501-3", "832Z4501-4", "832Z4501-5", "832Z4501-7", "832Z4501-6", "832Z4501-2",
            "832Z4501-8", "832Z4501-9", "832Z4501-11"]
    revA = ["830Z1009-2724", "830Z1009-2736", "830Z1009-2748", "830Z2042-407##ALT1", "830Z2042-408##ALT1"]
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
                        products_act_to_replace_nonc.ReplaceProduct(product_act_to_replace, Prod_replacing_part, True)

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


def ZZZ_AddNewIRMs():
    alldata = ZZZ_PartNumberCreator(SOW)
    pn = alldata[0]
    print pn
    id_new = alldata[1]
    print id_new
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products

    for item in range(len(pn)):
        product2 = products1.AddNewComponent("Product", pn[item])
        product2.name = id_new[item]


def reframe_all():

    """
    Ececutes 'Reframe All' CATIA command
    :return: nothing
    """

    ICM_1.ApplyWorkMode(2)
    specsAndGeomWindow1 = CATIA.ActiveWindow
    viewer3D1 = specsAndGeomWindow1.ActiveViewer
    viewer3D1.Reframe()


def ZZZ_PartNumberCreator(path1):
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
    return PNs, InstanceIDs


def irm_builder_nonconstant():

    """
    :return: no return. Walks the portion of ECS details located into the non-constant part of the airplane
    and calling corresponding functions to create all types of IRMs and CARMs
    """

    global make_carms
    global order_of_templete_product
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    first_elem_in_irm = []
    if make_carms_UI:
        make_carms = True

    for prod in xrange(1, 5):
        selection1 = productDocument1.Selection
        selection1.Clear()
        state = True
        product_inwork = products1.Item(prod)
        print product_inwork.PartNumber
        if '47' in product_inwork.PartNumber:
            tail_section = True
        else:
            tail_section = False
        products_inwork = product_inwork.Products

        # UPPERS NON-CONSTANT

        for det in xrange(1, products_inwork.Count+1):
            selection1 = productDocument1.Selection
            selection1.Clear()
            product_inwork_nonc = products_inwork.Item(det)
            print product_inwork_nonc.PartNumber
            name = product_inwork_nonc.Name
            print name
            if 'FAIRING' in str(product_inwork_nonc.PartNumber):
                if not state:
                    product_inwork_nonc = products_inwork.Item(det-1)
                    if not tail_section:
                        instantiate_carm_upper_final_nonc(product_inwork_nonc, name, product_forpaste_upr, first_elem_in_irm)
                        instantiate_carm_lwr_nonc(product_inwork_nonc, name, product_forpaste_lwr, 'final', first_elem_in_irm)
                    else:
                        instantiate_carm_upper_final_nonc_s47(product_inwork_nonc, name, product_forpaste_upr, first_elem_in_irm)
                        instantiate_carm_lwr_nonc_s47(product_inwork_nonc, name, product_forpaste_lwr, 'final', first_elem_in_irm)
                    first_elem_in_irm[:] = []
                product_inwork_nonc = products_inwork.Item(det)
                irm_type = 'OMF'
                if not tail_section:
                    instantiate_omf_nonconstant_irm_and_carm(product_inwork_nonc, name)
                else:
                    instantiate_omf_nonconstant_section47_irm_and_carm(product_inwork_nonc, name)
                state = True
            else:
                products_inwork_nonc = product_inwork_nonc.Products
                irm_type = 'UPR'
                if state:
                    add_new_irm(irm_type)
                    product_forpaste_upr = products1.Item('ECS_' + irm_type + '-AIR-DIST_INSTL_STA' + str(dash_number - 1))
                for det_deep in xrange(1, 3):
                    product_highl_inwork_nonc = products_inwork_nonc.Item(det_deep)
                    selection1.Add(product_highl_inwork_nonc)

                paste(selection1, product_forpaste_upr)
                first_elem_in_irm.append(name)
                print first_elem_in_irm

                # LOWERS NON-CONSTANT:

                selection1 = productDocument1.Selection
                selection1.Clear()
                irm_type = 'LWR'
                if state:
                    add_new_irm(irm_type)
                    product_forpaste_lwr = products1.Item('ECS_' + irm_type + '-AIR-DIST_INSTL_STA' + str(dash_number - 1))
                for det_deep in xrange(3, products_inwork_nonc.Count):
                    product_highl_inwork_nonc = products_inwork_nonc.Item(det_deep)
                    selection1.Add(product_highl_inwork_nonc)

                paste(selection1, product_forpaste_lwr)
                order_of_templete_product += 1
                if not tail_section:
                    if state:
                        instantiate_carm_upper_nonconstant(product_inwork_nonc, name, product_forpaste_upr)
                        instantiate_carm_lwr_nonc(product_inwork_nonc, name, product_forpaste_lwr, 'initial', first_elem_in_irm)
                    else:
                        instantiate_carm_upper_middle_nonconstant(product_inwork_nonc, name, product_forpaste_upr)
                        instantiate_carm_lwr_nonc(product_inwork_nonc, name, product_forpaste_lwr, 'middle', first_elem_in_irm)
                state = False
                if det is products_inwork.Count:
                    print 'FOUND IT'
                    if not tail_section:
                        instantiate_carm_upper_final_nonc(product_inwork_nonc, name, product_forpaste_upr, first_elem_in_irm)
                        instantiate_carm_lwr_nonc(product_inwork_nonc, name, product_forpaste_lwr, 'final', first_elem_in_irm)
                    else:
                        instantiate_carm_upper_final_nonc_s47(product_inwork_nonc, name, product_forpaste_upr, first_elem_in_irm)
                        instantiate_carm_lwr_nonc_s47(product_inwork_nonc, name, product_forpaste_lwr, 'final', first_elem_in_irm)
                    first_elem_in_irm[:] = []
                    state = True


def irm_builder_constant():

    """
    :return: no return. Walks the portion of ECS details located into the constant part of the airplane
    and calling corresponding functions to create all types of IRMs and CARMs
    """

    global constSize
    global order_of_templete_product
    global make_carms
    # -9 Bin breaks
    if plug == 240:
        bin_breaks = [561, 690 + 120, 897 + 120, 1089 + 120, 1290 + 120, 1401 + 96 + 120, 1560 + 240]
    # -8 Bin breaks
    elif plug == 0:
        bin_breaks = [690, 897, 1089, 1290, 1470, 1560]
    # -10 Bin breaks
    else:
        bin_breaks = [609 + 48, 690 + 240, 897 + 240, 1089 + 240, 1290 + 240, 1401 + 90 + 240, 1618 + 240 + 120 + 40]

    breaker = 0
    num = 0
    state = True
    initial_side = 'LH'
    switch_side = 0
    first_elem_in_irm = []
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products

    for prod in xrange(5, constSize+5):
        selection1 = productDocument1.Selection
        selection1.Clear()
        product_inwork = products1.Item(prod)
        print product_inwork.PartNumber
        name = product_inwork.Name
        print name
# UPPERS:

        if initial_side == 'LH':
                if 'RH' in str(product_inwork.Name):
                    initial_side = 'RH'
                    num = 0
                    state = True

        if 'FAIRING' in str(product_inwork.PartNumber):
            if not state:
                product_inwork = products1.Item(prod-1)
                instantiate_upper_bin_carm_final(product_inwork, name, product_forpaste_upr, first_elem_in_irm)
                instantiate_carm_lwr(product_inwork, name, product_forpaste_lwr, 'final', first_elem_in_irm)
                first_elem_in_irm[:] = []
            product_inwork = products1.Item(prod)
            irm_type = 'OMF'
            instantiate_omf_irm_and_carm(product_inwork, name)
            state = True

        else:
            products_inwork = product_inwork.Products
            make_carms = not if_stable()
            irm_type = 'UPR'
            if state:
                add_new_irm(irm_type)
                product_forpaste_upr = products1.Item('ECS_' + irm_type + '-AIR-DIST_INSTL_STA' + str(dash_number - 1))
            for det_deep in xrange(1, 3):
                product_highl_inwork = products_inwork.Item(det_deep)
                print product_highl_inwork.name
                selection1.Add(product_highl_inwork)

            paste(selection1, product_forpaste_upr)
            first_elem_in_irm.append(name)
            print first_elem_in_irm
# LOWERS:
            selection1 = productDocument1.Selection
            selection1.Clear()
            irm_type = 'LWR'
            if state:
                add_new_irm(irm_type)
                product_forpaste_lwr = products1.Item('ECS_' + irm_type + '-AIR-DIST_INSTL_STA' + str(dash_number - 1))
            for det_deep in xrange(3, products_inwork.Count):
                product_highl_inwork = products_inwork.Item(det_deep)
                selection1.Add(product_highl_inwork)

            paste(selection1, product_forpaste_lwr)
            order_of_templete_product += 1
            if state:
                instantiate_upper_bin_carm_initial(product_inwork, name, product_forpaste_upr)
                instantiate_carm_lwr(product_inwork, name, product_forpaste_lwr, 'initial', first_elem_in_irm)
            elif not state:
                instantiate_upper_bin_carm_middle(product_inwork, name, product_forpaste_upr)
                instantiate_carm_lwr(product_inwork, name, product_forpaste_lwr, 'middle', first_elem_in_irm)

        if bin_breaks[num] <= bin_breaker[breaker + 1]:
            state = True
            print first_elem_in_irm
            if irm_type != 'OMF':
                instantiate_upper_bin_carm_final(product_inwork, name, product_forpaste_upr, first_elem_in_irm)
                instantiate_carm_lwr(product_inwork, name, product_forpaste_lwr, 'final', first_elem_in_irm)
                print first_elem_in_irm
                first_elem_in_irm[:] = []
            if (num + 1) == len(bin_breaks):
                num = 0
                switch_side = 1
                
            else:
                num += 1
            if (breaker + 2) < len(bin_breaker):
                breaker += 1

        else:
            if irm_type != 'OMF':
                state = False
                if switch_side == 1:
                    state = True
                    print first_elem_in_irm
                    instantiate_upper_bin_carm_final(product_inwork, name, product_forpaste_upr, first_elem_in_irm)
                    instantiate_carm_lwr(product_inwork, name, product_forpaste_lwr, 'final', first_elem_in_irm)
                    print first_elem_in_irm
                    first_elem_in_irm[:] = []
                    switch_side = 0
            if (breaker + 2) < len(bin_breaker):
                breaker += 1
            continue


def add_new_irm(irm_type):

    """
    :param irm_type: 'omf', 'upr' or 'lwr'
    :return: adds new IRM product and returns it
    """

    global order_of_templete_product
    global order_of_new_product
    global dash_number
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    selection1 = productDocument1.Selection
    selection1.Clear()
    pn = 'IR830ZXXXX-' + str(dash_number)
    id = 'ECS_' + irm_type + '-AIR-DIST_INSTL_STA' + str(dash_number)
    product_forpaste = products1.AddNewComponent("Product", pn)
    product_forpaste.name = id
    order_of_new_product = dash_number - 995 + constSize
    print order_of_new_product
    dash_number += 1
    return product_forpaste


def paste(selection1, product_forpaste):

    """
    :param selection1: selected elements
    :param product_forpaste: product to paste
    :return: doesn't return anything. Pastes selected elements into the product
    """

    productDocument1 = CATIA.ActiveDocument
    selection1.Copy()
    selection1.Clear()
    selection2 = productDocument1.Selection
    selection2.Clear()
    selection2.add(product_forpaste)
    selection2.Paste()
    selection2.Clear()


def instantiate_omf_irm_and_carm(product_in_work, name):

    global order_of_templete_product
    global order_of_new_product
    global dash_number
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    selection1 = productDocument1.Selection
    selection1.Clear()
    size = product_in_work.Name[:2]
    start_sta = sta_values_fake[order_of_templete_product - 4]
    finish_sta = sta_value((sta_value_pairs[order_of_templete_product - 4] + inch_to_mm(int(size))), plug)
    pn = 'IR830ZXXXX-' + str(dash_number)
    if 'LH' in name:
        side_id = 'L'
    else:
        side_id = 'R'
    id = 'ECS_OMF_UPR-AIR_INSTL_STA' + start_sta + '-' + finish_sta + '_' + side_id
    product_forpaste = products1.AddNewComponent("Product", pn)
    product_forpaste.name = id
    order_of_new_product = dash_number - 995 + constSize
    print order_of_new_product
    dash_number += 1

    copy_from_name = product_in_work.Name
    print copy_from_name
    products_inwork_nonc = product_in_work.Products

    if 'ARCH' in product_in_work.Name[-7:]:
            arch = True
    else:
            arch = False

    if 'LH' in product_in_work.Name:
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
    order_of_templete_product += 1
    print order_of_new_product
    product_name = product_forpaste.name
    product_pn = product_forpaste.PartNumber
    carm_pn = product_pn[2:]
    carm_name = product_name + '_CARM'
    print carm_pn
    print carm_name

    if make_carms:

        carm_instance = CarmOmf(carm_pn, carm_name, side, order_of_new_product, order_of_templete_product, name)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        carm_instance.add_carm_as_external_component()
        carm_instance.change_inst_id()
        carm_instance.set_parameters(sta_value_pairs, size)
        carm_instance.modif_ref_annotation(size)
        carm_instance.modif_sta_annotation(sta_values_fake)
        carm_instance.copy_ref_surface_and_paste(size)
        carm_instance.copy_bodies_and_paste('BACS12FA3K3')
        carm_instance.copy_bodies_and_paste('FCM10F5CPS05WH')
        carm_instance.copy_jd1_fcm10f5cps05wh_and_paste(size)
        carm_instance.copy_jd2_bacs12fa3k3_and_paste(size, arch)
        carm_instance.set_standard_parts_params(1)
        carm_instance.set_standard_parts_params(2)
        carm_instance.create_jd_vectors(1)
        carm_instance.create_jd_vectors(2)
        carm_instance.access_captures(4)
        carm_instance.add_jd_annotation('01', sta_value_pairs, size, side)
        carm_instance.access_captures(5)
        carm_instance.add_jd_annotation('02', sta_value_pairs, size, side)
        carm_instance.shift_camera(sta_value_pairs, size)
        carm_instance.access_captures(1)
        carm_instance.hide_unhide_captures('unhide', 1)
        carm_instance.activate_top_prod()


def instantiate_omf_nonconstant_irm_and_carm(product_in_work_nonconstant, name):

    global order_of_templete_product
    global order_of_new_product
    global dash_number
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    selection1 = productDocument1.Selection
    selection1.Clear()
    start_sta = sta_values_fake[order_of_templete_product - 4]
    finish_sta = sta_values_fake[order_of_templete_product - 5]
    pn = 'IR830ZXXXX-' + str(dash_number)
    if 'LH' in name:
        side_id = 'L'
    else:
        side_id = 'R'
    id = 'ECS_OMF_UPR-AIR_INSTL_STA' + start_sta + '-' + finish_sta + '_' + side_id
    product_forpaste = products1.AddNewComponent("Product", pn)
    product_forpaste.name = id
    order_of_new_product = dash_number - 995 + constSize
    print order_of_new_product
    dash_number += 1

    copy_from_name = product_in_work_nonconstant.Name
    print copy_from_name
    products_inwork_nonc = product_in_work_nonconstant.Products

    size = product_in_work_nonconstant.Name[:2]
    if 'ARCH' in product_in_work_nonconstant.Name[-7:]:
            arch = True
    else:
            arch = False

    if 'LH' in product_in_work_nonconstant.Name:
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
    order_of_templete_product += 1
    print order_of_new_product
    product_name = product_forpaste.name
    product_pn = product_forpaste.PartNumber
    carm_pn = product_pn[2:]
    carm_name = product_name + '_CARM'
    print carm_pn
    print carm_name

    if make_carms:

        carm_instance = CarmUpperBinNonConstant(carm_pn, carm_name, side, order_of_new_product, order_of_templete_product, name)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        carm_instance.add_carm_as_external_component()
        carm_instance.change_inst_id()
        carm_instance.set_parameters(sta_value_pairs, size)
        carm_instance.modif_ref_annotation(size)
        carm_instance.modif_sta_annotation(sta_values_fake)
        carm_instance.copy_ref_surface_and_paste(size)
        carm_instance.copy_bodies_and_paste('BACS12FA3K3')
        carm_instance.copy_bodies_and_paste('FCM10F5CPS05WH')
        carm_instance.copy_jd1_fcm10f5cps05wh_and_paste(size)
        carm_instance.copy_jd2_bacs12fa3k3_and_paste(size, arch)
        carm_instance.set_standard_parts_params(1)
        carm_instance.set_standard_parts_params(2)
        carm_instance.create_jd_vectors(1)
        carm_instance.create_jd_vectors(2)
        carm_instance.access_captures(4)
        carm_instance.add_jd_annotation('01', sta_value_pairs, size, side, arch)
        carm_instance.access_captures(5)
        carm_instance.add_jd_annotation('02', sta_value_pairs, size, side, arch)
        carm_instance.shift_camera(sta_value_pairs, size)
        carm_instance.access_captures(1)
        carm_instance.hide_unhide_captures('unhide', 1)
        carm_instance.activate_top_prod()


def if_stable():

    """
    Checks if an IRM is in stable zone, if yes - CARM won't be populated. Currently works only for the 787-9
    :return: True if IRM is in stable zone, False if not
    """

    if make_carms_UI:
        if plug == 240:
            #stable_zone = ['0465', '0513', '0897', '0945', '0993', '1041', '1365', '1401', '1401+0', '1401+48', '1401+96', '1425', '1473', '1521', '1569']
            stable_zone = ['0897', '0945', '0993', '1041', '1365', '1401', '1401+0', '1401+48', '1401+96', '1425', '1473', '1521', '1569']
            sta = sta_values_fake[order_of_templete_product - 4]
            if sta in stable_zone:
                return True
            else:
                return False
        else:
            return False
    else:
        return True


def instantiate_omf_nonconstant_section47_irm_and_carm(product_in_work_nonconstant, name):
    global order_of_templete_product
    global order_of_new_product
    global dash_number
    productDocument1 = CATIA.ActiveDocument
    product1 = productDocument1.Product
    products1 = product1.Products
    selection1 = productDocument1.Selection
    selection1.Clear()
    size = product_in_work_nonconstant.Name[:2]
    start_sta = sta_values_fake[order_of_templete_product - 4]
    finish_sta = int(sta_values_fake[order_of_templete_product - 4]) + int(size)
    pn = 'IR830ZXXXX-' + str(dash_number)
    if 'LH' in name:
        side_id = 'L'
    else:
        side_id = 'R'
    id = 'ECS_OMF_UPR-AIR_INSTL_STA' + start_sta + '-' + str(finish_sta) + '_' + side_id
    product_forpaste = products1.AddNewComponent("Product", pn)
    product_forpaste.name = id
    order_of_new_product = dash_number - 995 + constSize
    print order_of_new_product
    dash_number += 1

    copy_from_name = product_in_work_nonconstant.Name
    print copy_from_name
    products_inwork_nonc = product_in_work_nonconstant.Products

    if 'ARCH' in product_in_work_nonconstant.Name[-7:]:
            arch = True
    else:
            arch = False

    if 'LH' in product_in_work_nonconstant.Name:
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
    order_of_templete_product += 1
    print order_of_new_product
    product_name = product_forpaste.name
    product_pn = product_forpaste.PartNumber
    carm_pn = product_pn[2:]
    carm_name = product_name + '_CARM'
    print carm_pn
    print carm_name


def instantiate_upper_bin_carm_initial(product_in_work, name, product_for_paste_upr):

    """
    :param product_in_work:
    :param name:
    :param product_for_paste_upr:
    :return: creates CARM for upper plenum installations in constant section, makes initial modifications to it
    """

    copy_from_name = product_in_work.Name
    print copy_from_name

    size = product_in_work.Name[:2]
    print product_in_work.Name[-7:]
    if 'ARCH' in product_in_work.Name[-7:]:
            arch = True
    else:
            arch = False
    if 'LH' in product_in_work.Name:
        side = 'LH'
    else:
        side = 'RH'
    product_name = product_for_paste_upr.name
    product_pn = product_for_paste_upr.PartNumber
    carm_pn = product_pn[2:]
    carm_name = product_name + '_CARM'
    print carm_pn
    print carm_name

    if make_carms:

        carm_instance = CarmUpperBin(carm_pn, carm_name, side, (order_of_new_product - 1), order_of_templete_product, name)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        carm_instance.add_carm_as_external_component()
        carm_instance.set_parameters(sta_value_pairs, size)
        carm_instance.modif_ref_annotation(size)
        carm_instance.modif_sta_annotation(sta_values_fake)
        carm_instance.copy_jd1_fcm10f5cps05wh_and_paste(size)
        carm_instance.copy_jd2_bacs12fa3k3_and_paste(size, arch)
        carm_instance.copy_ref_surface_and_paste(size)
        carm_instance.copy_bodies_and_paste('BACS12FA3K3')
        carm_instance.copy_bodies_and_paste('FCM10F5CPS05WH')
        carm_instance.shift_camera(sta_value_pairs, size)


def instantiate_upper_bin_carm_middle(product_in_work, name, product_for_paste_upr):

    copy_from_name = product_in_work.Name
    print copy_from_name

    size = product_in_work.Name[:2]
    print product_in_work.Name[-7:]
    if 'ARCH' in product_in_work.Name[-7:]:
            arch = True
    else:
            arch = False
    if 'LH' in product_in_work.Name:
        side = 'LH'
    else:
        side = 'RH'
    product_name = product_for_paste_upr.name
    product_pn = product_for_paste_upr.PartNumber
    carm_pn = product_pn[2:]
    carm_name = product_name + '_CARM'
    print carm_pn
    print carm_name

    if make_carms:

        carm_instance = CarmUpperBin(carm_pn, carm_name, side, (order_of_new_product - 1), order_of_templete_product, name)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        carm_instance.access_captures(3)
        carm_instance.add_ref_annotation(sta_value_pairs, size, side)
        carm_instance.add_sta_annotation(sta_value_pairs, sta_values_fake, size, side)
        carm_instance.access_captures(1)
        carm_instance.hide_unhide_captures('unhide', 1)
        carm_instance.activate_top_prod()
        carm_instance.copy_ref_surface_and_paste(size)
        carm_instance.copy_bodies_and_paste('BACS12FA3K3')
        carm_instance.copy_bodies_and_paste('FCM10F5CPS05WH')
        carm_instance.copy_jd1_fcm10f5cps05wh_and_paste(size)
        carm_instance.copy_jd2_bacs12fa3k3_and_paste(size, arch)


def instantiate_upper_bin_carm_final(product_in_work, name, product_for_paste_upr, first_elem_in_irm):

    copy_from_name = product_in_work.Name
    print copy_from_name

    size = product_in_work.Name[:2]
    if 'ARCH' in product_in_work.Name[-7:]:

            arch = True
    else:

            arch = False

    if 'LH' in product_in_work.Name:
        side = 'LH'
    else:
        side = 'RH'

    product_name = product_for_paste_upr.name
    product_pn = product_for_paste_upr.PartNumber
    carm_pn = product_pn[2:]
    carm_name = product_name + '_CARM'
    print carm_pn
    print carm_name

    if make_carms:

        carm_instance = CarmUpperBin(carm_pn, carm_name, side, (order_of_new_product - 1), order_of_templete_product, name, first_elem_in_irm, plug)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        carm_instance.change_inst_id_sta(sta_values_fake, sta_value_pairs, side, size)
        carm_instance.change_inst_id()
        carm_instance.set_standard_parts_params(1)
        carm_instance.set_standard_parts_params(2)
        carm_instance.create_jd_vectors(1)
        carm_instance.create_jd_vectors(2)
        carm_instance.access_captures(4)
        carm_instance.add_jd_annotation('01', sta_value_pairs, size, side, arch)
        carm_instance.access_captures(5)
        carm_instance.add_jd_annotation('02', sta_value_pairs, size, side, arch)
        carm_instance.access_captures(1)
        carm_instance.hide_unhide_captures('unhide', 1)
        carm_instance.activate_top_prod()
    else:
        carm_instance = CarmUpperBin(carm_pn, carm_name, side, (order_of_new_product - 1), order_of_templete_product, name, first_elem_in_irm, plug)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        carm_instance.change_inst_id_sta(sta_values_fake, sta_value_pairs, side, size)


def instantiate_carm_upper_nonconstant(product_inwork_nonc, name, product_forpaste_upr):

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
    product_name = product_forpaste_upr.name
    product_pn = product_forpaste_upr.PartNumber
    carm_pn = product_pn[2:]
    carm_name = product_name + '_CARM'
    print carm_pn
    print carm_name

    if make_carms:
        # INITIAL
        carm_instance = CarmOmfNonConstant(carm_pn, carm_name, side, (order_of_new_product - 1), order_of_templete_product, name)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        carm_instance.add_carm_as_external_component()
        carm_instance.set_parameters(sta_value_pairs, size)
        carm_instance.modif_ref_annotation(size)
        carm_instance.modif_sta_annotation(sta_values_fake)
        carm_instance.copy_jd1_fcm10f5cps05wh_and_paste(size)
        carm_instance.copy_jd2_bacs12fa3k3_and_paste(size, arch)
        carm_instance.copy_ref_surface_and_paste(size)
        carm_instance.copy_bodies_and_paste('BACS12FA3K3')
        carm_instance.copy_bodies_and_paste('FCM10F5CPS05WH')
        carm_instance.shift_camera(sta_value_pairs, size)


def instantiate_carm_upper_middle_nonconstant(product_inwork_nonc, name, product_forpaste_upr):

    copy_from_name = product_inwork_nonc.Name
    print copy_from_name

    size = product_inwork_nonc.Name[:2]
    if 'ARCH' in product_inwork_nonc.Name[-7:]:
            arch = True
    else:
            arch = False
    if 'LH' in product_inwork_nonc.Name:
        side = 'LH'
    else:
        side = 'RH'
    product_name = product_forpaste_upr.name
    product_pn = product_forpaste_upr.PartNumber
    carm_pn = product_pn[2:]
    carm_name = product_name + '_CARM'
    print carm_pn
    print carm_name

    if make_carms:

        carm_instance = CarmOmfNonConstant(carm_pn, carm_name, side, (order_of_new_product - 1), order_of_templete_product, name)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        carm_instance.access_captures(3)
        carm_instance.add_ref_annotation(sta_value_pairs, size, side)
        carm_instance.add_sta_annotation(sta_value_pairs, sta_values_fake, size, side)
        carm_instance.access_captures(1)
        carm_instance.hide_unhide_captures('unhide', 1)
        carm_instance.activate_top_prod()
        carm_instance.copy_ref_surface_and_paste(size)
        carm_instance.copy_bodies_and_paste('BACS12FA3K3')
        carm_instance.copy_bodies_and_paste('FCM10F5CPS05WH')
        carm_instance.copy_jd1_fcm10f5cps05wh_and_paste(size)
        carm_instance.copy_jd2_bacs12fa3k3_and_paste(size, arch)


def instantiate_carm_upper_final_nonc(product_inwork_nonc, name, product_forpaste_upr, first_elem_in_irm):

    global plug
    copy_from_name = product_inwork_nonc.Name
    print copy_from_name

    size = product_inwork_nonc.Name[:2]
    if 'ARCH' in product_inwork_nonc.Name[-7:]:
            arch = True
    else:
            arch = False

    if 'LH' in product_inwork_nonc.Name:
        side = 'LH'
    else:
        side = 'RH'

    product_name = product_forpaste_upr.name
    product_pn = product_forpaste_upr.PartNumber
    carm_pn = product_pn[2:]
    carm_name = product_name + '_CARM'
    print carm_pn
    print carm_name

    if make_carms:

        carm_instance = CarmOmfNonConstant(carm_pn, carm_name, side, (order_of_new_product - 1), order_of_templete_product, name, first_elem_in_irm)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        carm_instance.change_inst_id_sta(sta_values_fake, side)
        carm_instance.change_inst_id()
        carm_instance.set_standard_parts_params(1)
        carm_instance.set_standard_parts_params(2)
        carm_instance.create_jd_vectors(1)
        carm_instance.create_jd_vectors(2)
        carm_instance.access_captures(4)
        carm_instance.add_jd_annotation('01', sta_value_pairs, size, side, arch)
        carm_instance.access_captures(5)
        carm_instance.add_jd_annotation('02', sta_value_pairs, size, side, arch)
        carm_instance.access_captures(1)
        carm_instance.hide_unhide_captures('unhide', 1)
        carm_instance.activate_top_prod()
    else:
        carm_instance = CarmOmfNonConstant(carm_pn, carm_name, side, (order_of_new_product - 1), order_of_templete_product, name, first_elem_in_irm)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        carm_instance.change_inst_id_sta(sta_values_fake, side)


def instantiate_carm_upper_final_nonc_s47(product_inwork_nonc, name, product_forpaste_upr, first_elem_in_irm):

    global plug
    copy_from_name = product_inwork_nonc.Name
    print copy_from_name

    size = product_inwork_nonc.Name[:2]
    if 'ARCH' in product_inwork_nonc.Name[-7:]:
            arch = True
    else:
            arch = False

    if 'LH' in product_inwork_nonc.Name:
        side = 'LH'
    else:
        side = 'RH'

    product_name = product_forpaste_upr.name
    product_pn = product_forpaste_upr.PartNumber
    carm_pn = product_pn[2:]
    carm_name = product_name + '_CARM'
    print carm_pn
    print carm_name

    if make_carms or not make_carms:

        carm_instance = CarmUpperBinNonConstantSection47(carm_pn, carm_name, side, (order_of_new_product - 1), order_of_templete_product, name, first_elem_in_irm, plug)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        carm_instance.change_inst_id_sta(sta_values_fake, sta_value_pairs, side, size)


def instantiate_carm_lwr(product_inwork, name, product_forpaste_upr, state, first_elem_in_irm):

    """
    :param product_inwork:
    :param name:
    :param product_forpaste_upr:
    :param state: 'initial', 'middle' of 'final'
    :param first_elem_in_irm:
    :return: creates CARM for lower plenum installations in constant section
    """

    global plug
    global parameters_set
    copy_from_name = product_inwork.Name
    print copy_from_name
    size = product_inwork.Name[:2]

    if 'ARCH' in product_inwork.Name[-7:]:
            arch = True
    else:
            arch = False
    if 'LH' in product_inwork.Name:
        side = 'LH'
    else:
        side = 'RH'
    product_name = product_forpaste_upr.name
    product_pn = product_forpaste_upr.PartNumber
    carm_pn = product_pn[2:]
    carm_name = product_name + '_CARM'
    print carm_pn
    print carm_name

    if make_carms:
        # INITIAL
        carm_instance = CarmLowerBin(carm_pn, carm_name, side, order_of_new_product, order_of_templete_product, name, state, first_elem_in_irm, plug)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        if state == 'initial':
            carm_instance.add_carm_as_external_component()
            if size != '24':
                carm_instance.set_parameters(sta_value_pairs, size)
                print 'done setting parameters initial'
            carm_instance.shift_camera(sta_value_pairs, size)
        if state != 'final':
            carm_instance.copy_jd1_BACS12FA3K20_and_paste(size)
            carm_instance.copy_jd2_bacs12fa3k3_and_paste_1(size)
            if size != '24':
                # if size != '24' and size != '36':
                carm_instance.copy_jd3_BACS12FA3K12_and_paste(size)
                if '24' in first_elem_in_irm[0] and not parameters_set:
                    carm_instance.set_parameters(sta_value_pairs, size)
                    parameters_set = True
                    print 'done setting parameters middle'
            carm_instance.copy_jd4_bacs12fa3k3_and_paste_2(size)
            carm_instance.copy_ref_surface_and_paste(size)
            carm_instance.copy_bodies_and_paste('156')
            carm_instance.copy_bodies_and_paste('BACS12FA3K3.')
            if size != '24':
                # if size != '24' and size != '36':
                carm_instance.copy_bodies_and_paste('BACS12FA3K12')
            carm_instance.copy_bodies_and_paste('BACS12FA3K20')
            carm_instance.copy_bodies_and_paste('BACS38K2')
            carm_instance.access_captures(3)
            carm_instance.add_ref_annotation(sta_value_pairs, size, side)
            carm_instance.add_sta_annotation(sta_value_pairs, sta_values_fake, size, side)
            carm_instance.access_captures(1)
            carm_instance.hide_unhide_captures('unhide', 1)
            carm_instance.activate_top_prod()
        if state == 'final':
            if not '36' in first_elem_in_irm[0] and not '42' in first_elem_in_irm[0] and not '48' in first_elem_in_irm[0] and not parameters_set:
                carm_instance.set_parameters(sta_value_pairs, size)
                print 'done setting parameters final'
            carm_instance.change_inst_id_sta(sta_values_fake, sta_value_pairs, side, size)
            carm_instance.change_inst_id()
            carm_instance.set_standard_parts_params(1)
            carm_instance.set_standard_parts_params(2)
            carm_instance.set_standard_parts_params(3)
            carm_instance.set_standard_parts_params(4)
            carm_instance.modif_lwr_strap_annotation()
            carm_instance.modif_upr_strap_annotation()
            carm_instance.access_captures(6)
            carm_instance.add_jd_annotation('01', sta_value_pairs, size, side, arch)
            carm_instance.access_captures(7)
            carm_instance.add_jd_annotation('02', sta_value_pairs, size, side, arch)
            carm_instance.access_captures(8)
            carm_instance.add_jd_annotation('03', sta_value_pairs, size, side, arch)
            carm_instance.access_captures(9)
            carm_instance.add_jd_annotation('04', sta_value_pairs, size, side, arch)
            carm_instance.access_captures(1)
            carm_instance.hide_unhide_captures('unhide', 1)
            carm_instance.create_jd_vectors(1)
            carm_instance.create_jd_vectors(2)
            carm_instance.create_jd_vectors(3)
            carm_instance.create_jd_vectors(4)
            carm_instance.activate_top_prod()
    else:
        carm_instance = CarmLowerBin(carm_pn, carm_name, side, order_of_new_product, order_of_templete_product, name, state, first_elem_in_irm, plug)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        if state == 'final':
            carm_instance.change_inst_id_sta(sta_values_fake, sta_value_pairs, side, size)


def instantiate_carm_lwr_nonc(product_inwork_nonc, name, product_forpaste_upr, state, first_elem_in_irm):

    """

    :param product_inwork_nonc:
    :param name:
    :param product_forpaste_upr:
    :param state: 'initial', 'middle' of 'final'
    :param first_elem_in_irm:
    :return: creates CARM for lower plenum installations in section 41
    """

    copy_from_name = product_inwork_nonc.Name
    print copy_from_name
    size = product_inwork_nonc.Name[:2]

    if 'ARCH' in product_inwork_nonc.Name[-7:]:
            arch = True
    else:
            arch = False
    if 'LH' in product_inwork_nonc.Name:
        side = 'LH'
    else:
        side = 'RH'
    product_name = product_forpaste_upr.name
    product_pn = product_forpaste_upr.PartNumber
    carm_pn = product_pn[2:]
    carm_name = product_name + '_CARM'
    print carm_pn
    print carm_name

    if make_carms == True:
        # INITIAL
        carm_instance = CarmLowerBinNonConstant(carm_pn, carm_name, side, order_of_new_product, order_of_templete_product, name, state, first_elem_in_irm)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        if state == 'initial':
            carm_instance.add_carm_as_external_component()
            if size != '24':
                carm_instance.set_parameters(sta_value_pairs, size)
            carm_instance.shift_camera(sta_value_pairs, size)
        if state != 'final':
            carm_instance.copy_jd1_BACS12FA3K20_and_paste(size)
            carm_instance.copy_jd2_bacs12fa3k3_and_paste_1(size)
            if size != '24':
                # if size != '24' and size != '36':
                carm_instance.copy_jd3_BACS12FA3K12_and_paste(size)
                if '24' in first_elem_in_irm[0]:
                    carm_instance.set_parameters(sta_value_pairs, size)
            carm_instance.copy_jd4_bacs12fa3k3_and_paste_2(size)
            carm_instance.copy_ref_surface_and_paste(size)
            carm_instance.copy_bodies_and_paste('156')
            carm_instance.copy_bodies_and_paste('BACS12FA3K3.')
            if size != '24':
                # if size != '24' and size != '36':
                carm_instance.copy_bodies_and_paste('BACS12FA3K12')
            carm_instance.copy_bodies_and_paste('BACS12FA3K20')
            carm_instance.copy_bodies_and_paste('BACS38K2')
            carm_instance.access_captures(3)
            carm_instance.add_ref_annotation(sta_value_pairs, size, side)
            carm_instance.add_sta_annotation(sta_value_pairs, sta_values_fake, size, side)
            carm_instance.access_captures(1)
            carm_instance.hide_unhide_captures('unhide', 1)
            carm_instance.activate_top_prod()
        if state == 'final':
            if not '36' in first_elem_in_irm[0] and not '42' in first_elem_in_irm[0] and not '48' in first_elem_in_irm[0]:
                    carm_instance.set_parameters(sta_value_pairs, size)
            carm_instance.change_inst_id_sta(sta_values_fake, side)
            carm_instance.change_inst_id()
            carm_instance.set_standard_parts_params(1)
            carm_instance.set_standard_parts_params(2)
            carm_instance.set_standard_parts_params(3)
            carm_instance.set_standard_parts_params(4)
            carm_instance.modif_lwr_strap_annotation()
            carm_instance.modif_upr_strap_annotation()
            carm_instance.access_captures(6)
            carm_instance.add_jd_annotation('01', sta_value_pairs, size, side, arch)
            carm_instance.access_captures(7)
            carm_instance.add_jd_annotation('02', sta_value_pairs, size, side, arch)
            carm_instance.access_captures(8)
            carm_instance.add_jd_annotation('03', sta_value_pairs, size, side, arch)
            carm_instance.access_captures(9)
            carm_instance.add_jd_annotation('04', sta_value_pairs, size, side, arch)
            carm_instance.access_captures(1)
            carm_instance.hide_unhide_captures('unhide', 1)
            carm_instance.create_jd_vectors(1)
            carm_instance.create_jd_vectors(2)
            carm_instance.create_jd_vectors(3)
            carm_instance.create_jd_vectors(4)
            carm_instance.activate_top_prod()
    else:
        carm_instance = CarmLowerBinNonConstant(carm_pn, carm_name, side, order_of_new_product, order_of_templete_product, name, state, first_elem_in_irm)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        if state == 'final':
            carm_instance.change_inst_id_sta(sta_values_fake, side)


def instantiate_carm_lwr_nonc_s47(product_inwork_nonc, name, product_forpaste_upr, state, first_elem_in_irm):

    """
    :param product_inwork_nonc:
    :param name:
    :param product_forpaste_upr:
    :param state: 'initial', 'middle' of 'final'
    :param first_elem_in_irm:
    :return:  creates CARM for lower plenum installations in section 47
    """

    copy_from_name = product_inwork_nonc.Name
    print copy_from_name

    size = product_inwork_nonc.Name[:2]

    if 'ARCH' in product_inwork_nonc.Name[-7:]:
            arch = True
    else:
            arch = False
    if 'LH' in product_inwork_nonc.Name:
        side = 'LH'
    else:
        side = 'RH'
    product_name = product_forpaste_upr.name
    product_pn = product_forpaste_upr.PartNumber
    carm_pn = product_pn[2:]
    carm_name = product_name + '_CARM'
    print carm_pn
    print carm_name

    if make_carms:
        # INITIAL
        carm_instance = CarmLowerBinNonConstantSection47(carm_pn, carm_name, side, order_of_new_product, order_of_templete_product, name, state, first_elem_in_irm, plug)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        if state == 'initial':
            carm_instance.add_carm_as_external_component()
            carm_instance.set_parameters(sta_value_pairs, size)
            carm_instance.shift_camera(sta_value_pairs, size)
        if state != 'final':
            carm_instance.copy_jd1_BACS12FA3K20_and_paste(size)
            carm_instance.copy_jd2_bacs12fa3k3_and_paste_1(size)
            if size != '24':
                # if size != '24' and size != '36':
                carm_instance.copy_jd3_BACS12FA3K12_and_paste(size)
            carm_instance.copy_jd4_bacs12fa3k3_and_paste_2(size)
            carm_instance.copy_ref_surface_and_paste(size)
            carm_instance.copy_bodies_and_paste('156')
            carm_instance.copy_bodies_and_paste('BACS12FA3K3.')
            if size != '24':
                # if size != '24' and size != '36':
                carm_instance.copy_bodies_and_paste('BACS12FA3K12')
            carm_instance.copy_bodies_and_paste('BACS12FA3K20')
            carm_instance.copy_bodies_and_paste('BACS38K2')
            carm_instance.access_captures(3)
            carm_instance.add_ref_annotation(sta_value_pairs, size, side)
            carm_instance.add_sta_annotation(sta_value_pairs, sta_values_fake, size, side)
            carm_instance.access_captures(1)
            carm_instance.hide_unhide_captures('unhide', 1)
            carm_instance.activate_top_prod()
        if state == 'final':
            carm_instance.change_inst_id_sta(sta_values_fake, sta_value_pairs, side, size)
    else:
        carm_instance = CarmLowerBinNonConstantSection47(carm_pn, carm_name, side, order_of_new_product, order_of_templete_product, name, state, first_elem_in_irm, plug)
        wb = carm_instance.workbench_id()
        if wb != 'Assembly':
            carm_instance.activate_top_prod()
        if state == 'final':
            carm_instance.change_inst_id_sta(sta_values_fake, sta_value_pairs, side, size)


def part_number_collector():

    """
    :return: prints to the shell a list replacedDetails which contain all the part numbers of details currently loaded
    in CATIA session
    """

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


def deleter():

    """
    :return: no return, deletes all library data
    """

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


def instantiate_nonconstant_components():

    """
    Instantiates 4 products for non-constant sections
    :return: no return
    """

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


def add_component(s, side, section, location, plug_value):

    """
    :param s: a list containing information about bin run
    :param side: 'LH' or 'RH' side
    :param section: 'constant' or 'nonconstant'
    :param location: 'nose', 'middle' or 'tail'
    :param plug_value: insert plug variable
    :return: Doesn't return anything, builds ECS layout using CATIA objects library and sets instance IDs. Modifies
    globals: sta_value_pairs, sta_values_fake
    """

    extention = '.CATProduct'
    global sta_value_pairs
    global sta_values_fake
    x_coord = inch_to_mm(465)
    x_coord_nonconstant = inch_to_mm(0)
    fake_coord_nonconstant_41 = inch_to_mm(459)
    if plug_value == 240:
        fake_coord_nonconstant_47 = inch_to_mm(1863)
    elif plug_value == 456:
        fake_coord_nonconstant_47 = inch_to_mm(2079)
    elif plug_value == 0:
        fake_coord_nonconstant_47 = inch_to_mm(1623)

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
            x_coord = inch_to_mm(693 + door2_coord)
            index += 1
            continue

        else:

            Rotate5 = [0.996194698, -0.087155742, 0, 0.087155742, 0.996194698, 0, 0, 0, 1, inch_to_mm(466.61647022),
                       inch_to_mm(0.08471639), 0]
            Rotate185 = [-0.996194698, -0.087155742, 0, 0.087155742, -0.996194698, 0, 0, 0, 1, inch_to_mm(466.61647018),
                         inch_to_mm(-0.084716377), 0]
            Rotate_5 = [0.998512978, 0.054514501, 0, -0.054514501, 0.998512978, 0, 0, 0, 1,
                        inch_to_mm(1618.61663822 + plug_value), inch_to_mm(0.17865996), 0]
            Rotate_185 = [-0.998512978, 0.054514501, 0, -0.054514501, -0.998512978, 0, 0, 0, 1,
                          inch_to_mm(1618.61663822 + plug_value), inch_to_mm(-0.17865996), 0]

            print int(number)

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

                    elif int(number) == 18 or int(number) == 181:
                        number = '18'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Eighteen_arch_RH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 30 or int(number) == 301:
                        number = '30'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_arch_RH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 54 or int(number) == 541:
                        number = '54'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fifty_four_arch_RH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 60 or int(number) == 601:
                        number = '60'
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

                    # PREMIUM:

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

                    elif int(number) == 182 or int(number) == 1812 or int(number) == 1821:
                        number = '18'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Eighteen_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 302 or int(number) == 3012 or int(number) == 3021:
                        number = '30'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 362:
                        dow_type = 'DWNR_JOG-STRT'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_six_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 542 or int(number) == 5412 or int(number) == 5421:
                        number = '54'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fifty_four_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 602 or int(number) == 6012 or int(number) == 6021:
                        number = '60'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Sixty_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 2412 or int(number) == 2421:
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Twenty_four_fairing_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '24'

                    elif int(number) == 3612 or int(number) == 3621:
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_six_fairing_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '36'

                    elif int(number) == 4212 or int(number) == 4221:
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fourty_two_fairing_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '42'

                    elif int(number) == 4812 or int(number) == 4821:
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fourty_eight_fairing_arch_RH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '48'

                    else:
                        x_coord += inch_to_mm(int(number))
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

                    elif int(number) == 18 or int(number) == 181:
                        number = '18'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Eighteen_arch_LH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 30 or int(number) == 301:
                        number = '30'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_arch_LH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 54 or int(number) == 541:
                        number = '54'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fifty_four_arch_LH_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 60 or int(number) == 601:
                        number = '60'
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

                        # PREM:

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

                    elif int(number) == 182 or int(number) == 1812 or int(number) == 1821:
                        number = '18'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Eighteen_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 302 or int(number) == 3012 or int(number) == 3021:
                        number = '30'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 362:
                        dow_type = 'DWNR_JOG-STRT'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_six_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 542 or int(number) == 5412 or int(number) == 5421:
                        number = '54'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fifty_four_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 602 or int(number) == 6012 or int(number) == 6021:
                        number = '60'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Sixty_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)

                    elif int(number) == 2412 or int(number) == 2421:
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Twenty_four_fairing_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '24'

                    elif int(number) == 3612 or int(number) == 3621:
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Thirty_six_fairing_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '36'

                    elif int(number) == 4212 or int(number) == 4221:
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fourty_two_fairing_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '42'

                    elif int(number) == 4812 or int(number) == 4821:
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = path + '\Fourty_eight_fairing_arch_LH_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention
                        oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                        PartDoc = CATIA.Documents.Open(PartDocPath1)
                        number = '48'

                    else:
                        x_coord += inch_to_mm(int(number))
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
                    if sta_value(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'LH' or sta_value(
                            x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'LH':
                        PartDocPath = path + '\Fourty_eight_horseshoe_solids_LH'
                    elif sta_value(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'RH' or sta_value(
                            x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'RH':
                        PartDocPath = path + '\Fourty_eight_horseshoe_solids_RH'
                    else:
                        PartDocPath = path + '\Fourty_eight_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 12 or int(number) == 121:
                    number = '12'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Twelve_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 18 or int(number) == 181:
                    number = '18'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Eighteen_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 30 or int(number) == 301:
                    number = '30'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Thirty_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 54 or int(number) == 541:
                    number = '54'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fifty_four_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 60 or int(number) == 601:
                    number = '60'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Sixty_solids'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 72 or int(number) == 721:
                    number = '72'
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

                    # Premium plenums:

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
                    if index != (len(s) - 1) and side == 'RH' and s[index] == '72' or index != (len(s) - 1) and side == 'LH' and s[index - 2] == '72':
                        PartDocPath = path + '\Twenty_four_EXT_DR3_LH_solids_pr'
                        dow_type = 'DWNR_JOG-STRT'
                        ligval_ammount = 2
                    elif index != (len(s) - 1) and side == 'RH' and s[index - 2] == '72' or index != (len(s) - 1) and side == 'LH' and s[index] == '72':
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
                    if sta_value(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'LH' or sta_value(
                            x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'LH':
                        PartDocPath = path + '\Fourty_eight_horseshoe_solids_LH'
                    elif sta_value(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'RH' or sta_value(
                            x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'RH':
                        PartDocPath = path + '\Fourty_eight_horseshoe_solids_RH'
                    else:
                        PartDocPath = path + '\Fourty_eight_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 122 or int(number) == 1212 or int(number) == 1221:

                    number = '12'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Twelve_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 182 or int(number) == 1812 or int(number) == 1821:

                    number = '18'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Eighteen_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 302 or int(number) == 3012 or int(number) == 3021:

                    number = '30'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Thirty_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 542 or int(number) == 5412 or int(number) == 5421:

                    number = '54'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fifty_four_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 602 or int(number) == 6012 or int(number) == 6021:

                    number = '60'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Sixty_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 722 or int(number) == 7212 or int(number) == 7221:

                    number = '72'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Seventy_two_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 2412 or int(number) == 2421:

                    number = '24'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Twenty_four_fairing_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 3612 or int(number) == 3621:

                    number = '36'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Thirty_six_fairing_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 4212 or int(number) == 4221:

                    number = '42'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_two_fairing_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 4812 or int(number) == 4821:

                    number = '48'
                    nozzl_type = 'PREM'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Fourty_eight_fairing_solids_pr'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                else:
                    x_coord += inch_to_mm(int(number))
                    iteration += 1
                    index += 1
                    continue

                    # SECTION 41:

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

                    # SECTION 47:

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

                elif int(number) == 12 or int(number) == 121:
                    number = '12'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Twelve_solids_sec47'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 18 or int(number) == 181:
                    number = '18'
                    iteration += 1
                    index += 1
                    PartDocPath = path + '\Eighteen_solids_sec47'
                    PartDocPath1 = PartDocPath + str(iteration) + extention
                    oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = CATIA.Documents.Open(PartDocPath1)

                elif int(number) == 30 or int(number) == 301:
                    number = '30'
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
                    x_coord += inch_to_mm(int(number))
                    iteration += 1
                    index += 1
                    continue

            # for lower plenums:

            if stowbin is True or 'twenty_four':
                if number == '36':
                    L_PL_size1 = '36'
                elif number == '42':
                    L_PL_size1 = '42'
                elif number == '48':
                    L_PL_size1 = '48'
                else:
                    L_PL_size1 = '24'

            if section == 'constant':

                NewComponent = ICM_Products.AddExternalComponent(PartDoc)
                PartDoc.Close()
                oFileSys.DeleteFile(PartDocPath1)
                NewComponent.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                NewComponent.Name = str(number) + 'IN STA ' + sta_value(x_coord, plug_value) + ' ' + side + Arch
                trouble2 = mm_to_inch(x_coord)
                bin_breaker.append(int(trouble2))
                sta_values_fake.append(sta_value(x_coord, plug_value))
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
                    Lower_Plenum1.name = 'CONST_' + L_PL_size1 + 'LWPLEN_STA' + sta_value(x_coord, plug_value) + '_' + \
                                         side[0]
                    Lower_Downer1.name = dow_type + '_STA' + sta_value(x_coord, plug_value) + '_' + side[0] + '1'
                    Lower_Downer2.name = dow_type + '_STA' + sta_value(x_coord, plug_value) + '_' + side[0] + '2'



                elif stowbin == 'twenty_four':
                    if ligval_ammount == 2:
                        Lower_Plenum1 = RenamingTool.Products.Item(4)
                        Lower_Downer1 = RenamingTool.Products.Item(5)
                        Lower_Plenum1.name = 'CONST_' + L_PL_size1 + 'LWPLEN_STA' + sta_value(x_coord,
                                                                                              plug_value) + '_' + side[0]
                        Lower_Downer1.name = dow_type + '_STA' + sta_value(x_coord, plug_value) + '_' + side[0]

                    elif ligval_ammount == 1:
                        Lower_Plenum1 = RenamingTool.Products.Item(3)
                        Lower_Downer1 = RenamingTool.Products.Item(4)
                        Lower_Plenum1.name = 'CONST_' + L_PL_size1 + 'LWPLEN_STA' + sta_value(x_coord,
                                                                                              plug_value) + '_' + side[0]
                        Lower_Downer1.name = dow_type + '_STA' + sta_value(x_coord, plug_value) + '_' + side[0]

                PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + sta_value(x_coord, plug_value) + '_' + \
                                  side[0]
                if len(PlenumAssy.name) > 24:
                    PlenumAssy.name = str(number) + nozzl_type + 'ASSY_' + 'STA' + sta_value(x_coord, plug_value) + '_' + \
                                      side[0]
                if ligval_ammount == 1:
                    LING_VAL.name = 'OB_BIN_LIGVAL_STA' + sta_value(x_coord, plug_value) + '_' + side[0]
                    if len(LING_VAL.name) > 24:
                        LING_VAL.name = 'OB_LIGVAL_STA' + sta_value(x_coord, plug_value) + '_' + side[0]
                elif ligval_ammount == 2:
                    LING_VAL.name = 'OB_BIN_LIGVAL_STA' + sta_value(x_coord, plug_value) + '_' + side[0] + '1'
                    if len(LING_VAL.name) > 24:
                        LING_VAL.name = 'OB_LIGVAL_STA' + sta_value(x_coord, plug_value) + '_' + side[0] + '1'
                    LING_VAL2.name = 'OB_BIN_LIGVAL_STA' + sta_value(x_coord, plug_value) + '_' + side[0] + '2'
                    if len(LING_VAL2.name) > 24:
                        LING_VAL.name = 'OB_LIGVAL_STA' + sta_value(x_coord, plug_value) + '_' + side[0] + '2'

            elif section == 'nonconstant' and side == 'LH' and location == 'nose':

                NewComponent = ICM_Sec41_LH_Products.AddExternalComponent(PartDoc)
                PartDoc.Close()
                oFileSys.DeleteFile(PartDocPath1)
                RenamingToolProd = new_component1.ReferenceProduct
                Prod = RenamingToolProd.Products.Item(index)
                Prod.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                Prod.Name = str(number) + 'IN STA ' + sta_value(
                    (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))),
                    plug_value) + ' ' + side
                RenamingTool = NewComponent.ReferenceProduct
                PlenumAssy = RenamingTool.Products.Item(1)
                LING_VAL = RenamingTool.Products.Item(2)
                sta_values_fake.append(sta_value(
                    (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))),
                    plug_value))
                sta_value_pairs.append(x_coord_nonconstant)
                print sta_value_pairs
                print sta_values_fake

                if stowbin is True:
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Downer2 = RenamingTool.Products.Item(5)
                    Lower_Plenum1.name = 'SEC41_' + L_PL_size1 + 'LWPLEN_STA' + sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0]
                    Lower_Downer1.name = dow_type + '_STA' + sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0] + '1'
                    Lower_Downer2.name = dow_type + '_STA' + sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0] + '2'

                elif stowbin == 'twenty_four':
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Plenum1.name = 'SEC41_' + L_PL_size1 + 'LWPLEN_STA' + sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0]
                    Lower_Downer1.name = dow_type + '_STA' + sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0]

                PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + sta_value(
                    (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))), plug_value) + '_' + \
                                  side[0]
                LING_VAL.name = 'OB_BIN_LIGVAL_STA' + sta_value(
                    (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))), plug_value) + '_' + \
                                side[0]
                NewComponent.Move.Apply(Rotate5)

            elif section == 'nonconstant' and side == 'RH' and location == 'nose':

                NewComponent = ICM_Sec41_RH_Products.AddExternalComponent(PartDoc)
                PartDoc.Close()
                oFileSys.DeleteFile(PartDocPath1)
                RenamingToolProd = new_component2.ReferenceProduct
                Prod = RenamingToolProd.Products.Item(index)
                Prod.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                Prod.Name = str(number) + 'IN STA ' + sta_value(
                    (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))),
                    plug_value) + ' ' + side
                RenamingTool = NewComponent.ReferenceProduct
                PlenumAssy = RenamingTool.Products.Item(1)
                LING_VAL = RenamingTool.Products.Item(2)
                sta_values_fake.append(sta_value(
                    (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))),
                    plug_value))
                sta_value_pairs.append(x_coord_nonconstant)
                print sta_value_pairs
                print sta_values_fake

                if stowbin is True:
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Downer2 = RenamingTool.Products.Item(5)
                    Lower_Plenum1.name = 'SEC41_' + L_PL_size1 + 'LWPLEN_STA' + sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0]
                    Lower_Downer1.name = dow_type + '_STA' + sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0] + '1'
                    Lower_Downer2.name = dow_type + '_STA' + sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0] + '2'

                elif stowbin == 'twenty_four':
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Plenum1.name = 'SEC41_' + L_PL_size1 + 'LWPLEN_STA' + sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0]
                    Lower_Downer1.name = dow_type + '_STA' + sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))), plug_value) + '_' + \
                                         side[0]

                PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + sta_value(
                    (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))), plug_value) + '_' + \
                                  side[0]
                LING_VAL.name = 'OB_BIN_LIGVAL_STA' + sta_value(
                    (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))), plug_value) + '_' + \
                                side[0]
                NewComponent.Move.Apply(Rotate185)

            elif section == 'nonconstant' and side == 'LH' and location == 'tail':

                NewComponent = ICM_Sec47_LH_Products.AddExternalComponent(PartDoc)
                PartDoc.Close()
                oFileSys.DeleteFile(PartDocPath1)
                RenamingToolProd = new_component3.ReferenceProduct
                Prod = RenamingToolProd.Products.Item(index)
                Prod.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                Prod.Name = str(number) + 'IN STA ' + sta_value((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                plug_value) + ' ' + side
                RenamingTool = NewComponent.ReferenceProduct
                PlenumAssy = RenamingTool.Products.Item(1)
                Felt = RenamingTool.Products.Item(2)

                sta_values_fake.append(sta_value(
                    (fake_coord_nonconstant_47 + x_coord_nonconstant),
                    plug_value))
                sta_value_pairs.append(x_coord_nonconstant)
                print sta_value_pairs
                print sta_values_fake

                if stowbin is True:
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Downer2 = RenamingTool.Products.Item(5)
                    Lower_Plenum1.name = 'CONST_' + L_PL_size1 + 'LWPLEN_STA' + sta_value(
                        (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                    Lower_Downer1.name = dow_type + '_STA' + sta_value((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                       plug_value) + '_' + side[0] + '1'
                    Lower_Downer2.name = dow_type + '_STA' + sta_value((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                       plug_value) + '_' + side[0] + '2'

                elif stowbin == 'twenty_four':
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Plenum1.name = 'CONST_' + L_PL_size1 + 'LWPLEN_STA' + sta_value(
                        (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                    Lower_Downer1.name = dow_type + '_STA' + sta_value((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                       plug_value) + '_' + side[0]

                PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + sta_value(
                    (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                Felt.name = 'UPR_FELT_' + str(number) + 'IN_STA' + sta_value(
                    (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                NewComponent.Move.Apply(Rotate_5)

            elif section == 'nonconstant' and side == 'RH' and location == 'tail':

                NewComponent = ICM_Sec47_RH_Products.AddExternalComponent(PartDoc)
                PartDoc.Close()
                oFileSys.DeleteFile(PartDocPath1)
                RenamingToolProd = new_component4.ReferenceProduct
                Prod = RenamingToolProd.Products.Item(index)
                Prod.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                Prod.Name = str(number) + 'IN STA ' + sta_value((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                plug_value) + ' ' + side
                RenamingTool = NewComponent.ReferenceProduct
                PlenumAssy = RenamingTool.Products.Item(1)
                Felt = RenamingTool.Products.Item(2)

                sta_values_fake.append(sta_value(
                    (fake_coord_nonconstant_47 + x_coord_nonconstant),
                    plug_value))
                sta_value_pairs.append(x_coord_nonconstant)
                print sta_value_pairs
                print sta_values_fake

                if stowbin is True:
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Downer2 = RenamingTool.Products.Item(5)
                    Lower_Plenum1.name = 'CONST_' + L_PL_size1 + 'LWPLEN_STA' + sta_value(
                        (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                    Lower_Downer1.name = dow_type + '_STA' + sta_value((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                       plug_value) + '_' + side[0] + '1'
                    Lower_Downer2.name = dow_type + '_STA' + sta_value((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                       plug_value) + '_' + side[0] + '2'



                elif stowbin == 'twenty_four':
                    Lower_Plenum1 = RenamingTool.Products.Item(3)
                    Lower_Downer1 = RenamingTool.Products.Item(4)
                    Lower_Plenum1.name = 'CONST_' + L_PL_size1 + 'LWPLEN_STA' + sta_value(
                        (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                    Lower_Downer1.name = dow_type + '_STA' + sta_value((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                       plug_value) + '_' + side[0]

                PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + sta_value(
                    (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                Felt.name = 'UPR_FELT_' + str(number) + 'IN_STA' + sta_value(
                    (fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
                NewComponent.Move.Apply(Rotate_185)

            if location == 'nose':
                x_coord_nonconstant -= inch_to_mm(int(number))

            x = x_coord_nonconstant * math.cos(rad)
            y = x_coord_nonconstant * math.sin(rad)

            position = [1, 0, 0, 0, 1, 0, 0, 0, 1, x_coord, 0, 0]
            position_non = [1, 0, 0, 0, 1, 0, 0, 0, 1, x, -y, 0]
            position_non_RH = [1, 0, 0, 0, 1, 0, 0, 0, 1, x + (inch_to_mm(int(number)) * math.cos(rad)),
                               y + (inch_to_mm(int(number)) * math.sin(rad)), 0]
            position90 = [-1, 0, 0, 0, -1, 0, 0, 0, 1, x_coord + inch_to_mm(int(number)), 0, 0]  # 90 deg rotation
            position_non_47 = [1, 0, 0, 0, 1, 0, 0, 0, 1, x, y, 0]
            position_non_47_RH = [1, 0, 0, 0, 1, 0, 0, 0, 1, x + (inch_to_mm(int(number)) * math.cos(rad)),
                                  (y + (inch_to_mm(int(number)) * math.sin(rad))) * (-1), 0]

            if side == 'LH' and section == 'constant':
                # NewComponentRef = NewComponent.ReferenceProduct
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
                x_coord_nonconstant += inch_to_mm(int(number))
                print section
                print x, y
                print x_coord_nonconstant
            elif section == 'nonconstant' and side == 'RH' and location == 'tail':
                NewComponent.Move.Apply(position_non_47_RH)
                x_coord_nonconstant += inch_to_mm(int(number))
                print section
                print x, y
                print x_coord_nonconstant

            x_coord += inch_to_mm(int(number))


if __name__ == '__main__':
    root = Tkinter.Tk()
    TkFileDialogExample(root).pack()
    root.mainloop()
