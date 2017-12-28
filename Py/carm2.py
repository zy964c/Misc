import Tkinter
import Tkconstants
import tkFileDialog
import tkMessageBox
import math
import time
from APIE_Tool2v7_newadd import Inch_to_mm
import win32com.client


class CARM(object):

    def __init__(self, carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name):
        self.carm_part_number = carm_part_number
        self.order_of_new_product = order_of_new_product
        self.instance_id = instance_id
        self.copy_from_product = copy_from_product
        self.cpf_name = cfp_name
        self.catia = win32com.client.Dispatch('catia.application')
        self.path = ['C:\Users\zy964c\PycharmProjects']
        self.side = side
        self.extention = '\seed_carm_' + side + '.CATPart'
        self.oFileSys = self.catia.FileSystem
        self.productDocument1 = self.catia.ActiveDocument
        self.documents = self.catia.Documents


    def unused_instantiate_carm_new(self):

        productDocument1 = self.catia.ActiveDocument
        product1 = productDocument1.Product
        collection_irms = product1.Products
        product_to_insert_carm = collection_irms.Item(self.order_of_new_product)
        children_of_product_to_insert_carm = product_to_insert_carm.Products
        children_of_product_to_insert_carm.AddNewComponent("Part", self.carm_part_number)
        carm = children_of_product_to_insert_carm.Item(1)
        carm.ActivateDefaultShape()
        documents = self.catia.Documents
        self.carm_document = documents.Item(self.carm_part_number + '.CATPart')
    def unused_instantiate_carm_new_master_shape(self):

        productDocument1 = self.catia.ActiveDocument
        product1 = productDocument1.Product
        collection_irms = product1.Products
        product_to_insert_carm = collection_irms.Item(self.order_of_new_product)
        children_of_product_to_insert_carm = product_to_insert_carm.Products
        children_of_product_to_insert_carm.AddNewComponent("Product", self.carm_part_number)
        carm = children_of_product_to_insert_carm.Item(1)
        carm.AddMasterShapeRepresentation(self.path[0] + self.extention)
        carm.ActivateDefaultShape()
        documents = self.catia.Documents
        self.carm_document = documents.Item(self.carm_part_number + '.CATPart')
    def unused_instantiate_carm_from(self):

        productDocument1 = self.catia.ActiveDocument
        product1 = productDocument1.Product
        collection_irms = product1.Products
        product_to_insert_carm = collection_irms.Item(self.order_of_new_product)
        children_of_product_to_insert_carm = product_to_insert_carm.Products
        children_of_product_to_insert_carm.AddComponentsFromFiles(self.path + self.extention, "All")
        NewComponent = children_of_product_to_insert_carm.Item(1)
        NewComponent.ActivateDefaultShape()
        product1.ApplyWorkMode(2)
        NewComponent.PartNumber = self.carm_part_number
        NewComponent.Name = self.carm_part_number
        documents = self.catia.Documents
        self.carm_document = documents.Item(self.carm_part_number + self.extention)
    def unused_external_staff(self):

        productDocument1 = self.catia.ActiveDocument
        product1 = productDocument1.Product
        collection_irms = product1.Products
        product_to_insert_carm = collection_irms.Item(self.order_of_new_product)
        children_of_product_to_insert_carm = product_to_insert_carm.Products
        PartDocPath = self.path[0] + self.extention
        PartDocPath1 = self.path[0] + '\CA' + self.carm_part_number + '.CATPart'
        self.oFileSys.CopyFile(PartDocPath, PartDocPath1, False)
        PartDoc = self.catia.Documents.Open(PartDocPath1)
        children_of_product_to_insert_carm.AddExternalComponent(PartDoc)
        self.NewComponent = children_of_product_to_insert_carm.Item(1)
        PartDoc.Close()
        self.oFileSys.DeleteFile(PartDocPath1)
        self.NewComponent.PartNumber = self.carm_part_number
        self.catia.ActiveWindow.ActiveViewer.Reframe()
        self.documents = self.catia.Documents
        self.carm_document = self.documents.Item('CA' + self.carm_part_number + '.CATPart')


    def add_carm_as_external_component(self):
        """Instantiates CARM from external library"""

        product1 = self.productDocument1.Product
        collection_irms = product1.Products
        product_to_insert_carm = collection_irms.Item(self.order_of_new_product)
        children_of_product_to_insert_carm = product_to_insert_carm.Products
        PartDocPath = self.path[0] + self.extention
        PartDocPath1 = self.path[0] + '\CA' + self.carm_part_number + '.CATPart'
        self.oFileSys.CopyFile(PartDocPath, PartDocPath1, True)
        PartDoc = self.catia.Documents.NewFrom(PartDocPath1)
        PartDoc1 = PartDoc.Product
        PartDoc1.PartNumber = 'CA' + self.carm_part_number
        print PartDoc1.Name
        NewComponent = children_of_product_to_insert_carm.AddExternalComponent(PartDoc)
        #NewComponent = children_of_product_to_insert_carm.Item(1)
        PartDoc.Close()
        self.oFileSys.DeleteFile(PartDocPath1)
        print self.instance_id
        NewComponent.Name = self.instance_id
        print NewComponent.Name
        self.catia.ActiveWindow.ActiveViewer.Reframe()


    def change_inst_id(self):

        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        Product2 = to_p.ReferenceProduct
        Product2Products = Product2.Products
        product_forpaste = Product2Products.Item(Product2Products.Count)
        product_forpaste.Name = self.instance_id
        print product_forpaste.Name

    def access_carm(self):
        """Returns self carm_part"""

        carm_document = self.documents.Item('CA' + self.carm_part_number + '.CATPart')
        carm_part = carm_document.Part
        return carm_part

    def unused_access_to_part(self, pn):
        """Returns the part object"""

        documents1 = self.catia.Documents
        partDocument1 = documents1.Item(pn)
        part1 = partDocument1.Part
        return part1

    def unused_access_to_product(self, pn):
        """Returns the Product object"""

        documents1 = self.catia.Documents
        partDocument1 = documents1.Item(pn)
        part1 = partDocument1.Product
        return part1

    def access_annotations(self):

        carm_part = self.access_carm()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        anns = ann_set1.Annotations
        for ann in xrange(1, anns.Count+1):
            ann1 = anns.Item(ann)
            ann1text = ann1.Text()
            ann1text_2d = ann1text.Get2dAnnot()
            ann1text_value = ann1text_2d.Text
            print ann1text_value

    def modif_ref_annotation(self, size):

        carm_part = self.access_carm()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        anns = ann_set1.Annotations
        ref_annotation = anns.Item(1)
        ann1text = ref_annotation.Text()
        ann1text_2d = ann1text.Get2dAnnot()
        ann1text_value = size + 'IN OUTBD FRNG SUPPORT REF'
        ann1text_2d.Text = ann1text_value
        print ann1text_value

    def modif_sta_annotation(self, sta_values_fake):

        carm_part = self.access_carm()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        anns = ann_set1.Annotations
        sta_annotation = anns.Item(2)
        ann1text = sta_annotation.Text()
        ann1text_2d = ann1text.Get2dAnnot()
        sta = sta_values_fake[self.copy_from_product - 5]
        #ann1text_value = 'STA ' + sta + '\nLBL 74.3\nWL 294.8\nREF'
        ann1text_value = 'STA ' + sta + '\nREF'
        ann1text_2d.Text = ann1text_value
        print ann1text_value

    #def select_current_product(self):

        #"""Returns current product"""

        #Prod = self.productDocument1.Product
        #collection = Prod.Products
        #selection1.Search(str('Name =' + self.cpf_name + ', all'))
        #product1 = collection.Item(self.cpf_name)
        #print product1.name
        #return product1

    def select_current_product(self):
        #ICM_1.ApplyWorkMode(2)
        product1 = self.productDocument1.Product
        products1 = product1.Products

        for prod in xrange(1, 5):
            product_to_replace = products1.Item(prod)
            products_to_replace = product_to_replace.Products

            for det in xrange(1, products_to_replace.Count+1):
                product_act_to_replace_nonc = products_to_replace.Item(det)
                if self.cpf_name in str(product_act_to_replace_nonc.Name):
                    return product_act_to_replace_nonc
                else:
                    continue

        for prod in xrange(5, products1.Count+1):
            product_to_replace = products1.Item(prod)
            if self.cpf_name in str(product_to_replace.Name):
                return product_to_replace
            else:
                continue


    def select_carm_to_paste_data(self):
        """Returns part of the CARM through the reference product"""

        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        Product2 = to_p.ReferenceProduct
        Product2Products = Product2.Products
        product_forpaste = Product2Products.Item(Product2Products.Count)
        print product_forpaste.name
        Part3 = product_forpaste.ReferenceProduct
        PartDocument3 = Part3.Parent
        print PartDocument3.name
        geom_elem3 = PartDocument3.Part
        return geom_elem3

    def add_geosets(self):
        """Adds Reference Geometry geometrical set"""


        carm_part = self.access_carm()
        geosets = carm_part.HybridBodies
        new_geoset = geosets.add()
        new_geoset.name = 'Reference Geometry1'
        first_gs = geosets.Item(1)
        first_gs.name = 'Renamed'

        #for reference: For IdxSet = 1 To AnnotationSets.Count

    def get_points(self, jd_number):

        ref_connecter = []
        carm_part = self.access_carm()
        geosets = carm_part.HybridBodies
        geoset1 = geosets.Item('Joint Definitions')
        print geoset1.name
        geosets1 = geoset1.HybridBodies
        print 'Joint Definition ' + str(jd_number)
        geoset2 = geosets1.Item('Joint Definition ' + jd_number)
        print geoset2.name
        points = geoset2.HybridShapes
        for point in xrange(1, points.Count+1):
            target = points.Item(point)
            if not 'FIDV' in target.Name:
                ref_connecter.append(target.Name)
            else:
                continue
        print ref_connecter
        return ref_connecter
        #print ref_connecter.Name
        #ref_connecter_coordinates_X = ref_connecter.X
        #ref_connecter_coordinates_Y = ref_connecter.Y
        #ref_connecter_coordinates_Z = ref_connecter.Z
        #ref_connecter_new_coordinates.append(ref_connecter_coordinates_X.Value)
        #ref_connecter_new_coordinates.append(ref_connecter_coordinates_Y.Value)
        #ref_connecter_new_coordinates.append(ref_connecter_coordinates_Z.Value)
        #print ref_connecter_new_coordinates
        #ref_connecter_new_coordinates[0] += 300.0
        #print ref_connecter_new_coordinates
        #ref_connecter.SetCoordinates(ref_connecter_new_coordinates)

    def set_parameters(self, sta_value_pairs, size):

        carm_part = self.select_carm_to_paste_data()
        parameters1 = carm_part.Parameters
        ref_param = parameters1.Item('ref_connector_X')
        sta_param = parameters1.Item('sta_connector_X')
        #direct_param = parameters1.Item('view_direction_connector_X')
        print ref_param.Value
        print sta_param.Value
        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        print coord_to_move
        ref_param.Value = coord_to_move + (Inch_to_mm(float(size))) - Inch_to_mm(0.25)
        sta_param.Value = coord_to_move + Inch_to_mm(0.25)
        #direct_param.Value = coord_to_move + (Inch_to_mm(float(size)/2)) + Inch_to_mm(7.0)
        print ref_param.Value
        print sta_param.Value

    def set_standard_parts_params(self, jd_number):

        carm_part = self.select_carm_to_paste_data()
        hole_qty = 0
        parameters1 = carm_part.Parameters
        #for param1 in xrange(1, parameters1.Count):
            #param2 = parameters1.Item(param1)
            #print param2.Name
        selection1 = self.productDocument1.Selection
        selection1.Clear()
        hybridBodies1 = carm_part.HybridBodies
        hybridBody1 = hybridBodies1.Item("Joint Definitions")
        hybridBodies2 = hybridBody1.HybridBodies
        hybridBody2 = hybridBodies2.Item("Joint Definition" + ' 0' + str(jd_number))
        HybridShapes1 = hybridBody2.HybridShapes
        for shape in xrange(HybridShapes1.Count):
            hole_qty += 1
        print hole_qty
        param_hole_qty = parameters1.Item('Joint Definitions\Joint Definition 0' + str(jd_number) + '\Hole Quantity')
        param_hole_qty.Value = str(hole_qty)
        if jd_number == 1:
            param = parameters1.Item('FCM10F5CPS05WH')
            param.Value = str(hole_qty) + '|FCM10F5CPS05WH | 302 CRES, Passivated, with Captive Washer, Head Color'
        elif jd_number == 2:
            param = parameters1.Item('BACS12FA3K3')
            param.Value = str(hole_qty) + '|BACS12FA3K3 | SCREW, WASHER HEAD, CROSS RECESS, FULL THREADED, 6AL-4V TITANIUM'


    def unused_copy_bodies_and_paste(self, copy_from):
        """Makes copy of fasteners solids and pastes to current CARM"""

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        part1 = self.unused_access_to_part(copy_from)
        bodies1 = part1.Bodies
        body1 = bodies1.Item("BACS12FA3K3 REF")
        selection1.Add(body1)
        body2 = bodies1.Item("FCM10F5CPS05WH REF")
        selection1.Add(body2)
        selection1.Copy()
        selection2 = self.productDocument1.Selection
        selection2.Clear()
        part2 = self.access_carm()
        selection2.Add(part2)
        #selection2.PasteSpecial('CATPrtResultWithOutLink')
        selection2.Paste()
        part2.Update()

    def copy_bodies_and_paste(self, fastener):
        """Makes copy of fasteners solids and pastes them to the current CARM"""

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.select_current_product()
        selection1.Add(product1)
        selection1.Search(str('Name = ' + fastener + '*REF, sel'))
        selection1.Copy()
        selection2 = self.productDocument1.Selection
        selection2.Clear()
        part2 = self.select_carm_to_paste_data()
        selection2.Add(part2)
        selection2.PasteSpecial('CATPrtResultWithOutLink')
        part2.Update()

    def rename_part_body(self):
        """renames part body and activates it"""
        carm = self.access_carm()
        bodies = carm.Bodies
        part_body = bodies.Item(1)
        part_body.name = 'CA' + self.carm_part_number
        carm.InWorkObject = part_body

    def unused_copy_ref_surface_and_paste(self):
        """Makes copy of reference geometry geoset and pastes to current CARM"""

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.select_current_product()
        selection1.Add(product1)
        selection1.Search('Name = Reference Geometry, sel')
        selection1.Copy()
        selection2 = self.productDocument1.Selection
        selection2.Clear()
        geoset3 = self.select_carm_to_paste_data()
        selection2.Add(geoset3)
        selection2.visProperties.SetRealColor(0, 128, 255, 0)
        selection2.visProperties.SetRealOpacity(65, 0)
        selection2.visProperties.SetShow(0)
        selection2.PasteSpecial('CATPrtResultWithOutLink')
        geoset3.Update()
        #change visual properties
        selection3 = self.productDocument1.Selection
        selection3.Clear()
        geoset4 = self.select_carm_to_paste_data()
        selection3.Add(geoset4)
        selection3.Search('Name = Reference Geometry, sel')
        selection3.visProperties.SetRealColor(0, 128, 255, 0)
        selection3.visProperties.SetRealOpacity(65, 0)

    def copy_ref_surface_and_paste(self, size):
        """Makes copy of reference geometry geoset and pastes to current CARM"""

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.select_current_product()
        selection1.Add(product1)
        selection1.Search(str('Name = *' + size + 'IN*REF, sel'))
        selection1.Copy()
        selection2 = self.productDocument1.Selection
        selection2.Clear()
        geoset3 = self.select_carm_to_paste_data()
        hybridBodies1 = geoset3.HybridBodies
        hybridBody1 = hybridBodies1.Item("Reference Geometry")
        selection2.Add(hybridBody1)
        #selection2.visProperties.SetShow(0)
        selection2.PasteSpecial('CATPrtResultWithOutLink')
        geoset3.Update()
        #change visual properties
        selection3 = self.productDocument1.Selection
        selection3.Clear()
        geoset4 = self.select_carm_to_paste_data()
        selection3.Add(geoset4)
        selection3.Search('Name = Reference Geometry, sel')
        selection3.visProperties.SetRealColor(0, 128, 255, 0)
        selection3.visProperties.SetRealOpacity(65, 0)


    def unused_copy_jd2_bacs12fa3k3_and_paste(self, copy_from, size, arch):
        """Makes copy of points for JD2 fasteners and pastes to the current CARM"""

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        part1 = self.unused_access_to_part(copy_from)
        hybridBodies1 = part1.HybridBodies
        hybridBody1 = hybridBodies1.Item('ECS_FAIRING_INTERFACE_' + size + 'IN')
        hybridBodies2 = hybridBody1.HybridBodies
        if arch:
            pos_description = ' ARCH'
        else:
            pos_description = ''
        print size + 'IN PLENUMS - UPPER' + pos_description
        hybridBody2 = hybridBodies2.Item(size + 'IN PLENUMS - UPPER' + pos_description)
        hybridShapes1 = hybridBody2.HybridShapes
        for Plenum_spud in range(1, hybridShapes1.Count + 1):
            hybridShapeIntersection = hybridShapes1.Item(Plenum_spud)
            if not 'CENTERLINE' in hybridShapeIntersection.name and not 'ECS' in hybridShapeIntersection.name:
                selection1.Add(hybridShapeIntersection)
            else:
                continue
        selection1.Copy()
        self.paste_to_jd(2)

    def copy_jd2_bacs12fa3k3_and_paste(self, size, arch, type_of_geometry='points'):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.select_current_product()
        selection1.Add(product1)
        if type_of_geometry == 'points':
            if arch:
                selection1.Search(str('(Name = ' + size + '*FAIRING*PLENUM*ARCH*-Name = *CENTERLINE*), sel'))

            else:
                selection1.Search(str('(Name = ' + size + '*FAIRING*PLENUM*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))

        else:
            if arch:
                selection1.Search(str('(Name = ' + size + '*FAIRING*PLENUM*ARCH*1_CENTERLINE*),sel'))

            else:
                selection1.Search(str('(Name = ' + size + '*FAIRING*PLENUM*1_CENTERLINE*-(Name = *ARCH*)), sel'))
        selection1.Copy()
        self.paste_to_jd(2)

    def rename_vectors(self, jd_number):

        part2 = self.select_carm_to_paste_data()
        hybridBodies1 = part2.HybridBodies
        hybridBody1 = hybridBodies1.Item("Joint Definitions")
        hybridBodies2 = hybridBody1.HybridBodies
        hybridBody2 = hybridBodies2.Item("Joint Definition" + ' 0' + str(jd_number))
        hybridShapes1 = hybridBody2.HybridShapes
        for Plenum_spud in range(1, hybridShapes1.Count + 1):
            Vector= hybridShapes1.Item(Plenum_spud)
            if 'CENTERLINE' in Vector.name:
                Vector.name = 'FIDV_01'
            else:
                continue


    def unused_copy_jd1_fcm10f5cps05wh_and_paste(self, copy_from, size):
        """Makes copy of points for JD1 fasteners and pastes to the current CARM"""

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        part1 = self.unused_access_to_part(copy_from)
        hybridBodies1 = part1.HybridBodies
        hybridBody1 = hybridBodies1.Item('ECS_FAIRING_INTERFACE_' + size + 'IN')
        hybridBodies2 = hybridBody1.HybridBodies
        hybridBody2 = hybridBodies2.Item(size + 'IN LIGHT VALENCE')
        hybridShapes1 = hybridBody2.HybridShapes
        for light_valence in range(1, hybridShapes1.Count + 1):
            hybridShapeIntersection = hybridShapes1.Item(light_valence)
            if not 'CENTERLINE' in hybridShapeIntersection.name and not 'ECS' in hybridShapeIntersection.name:
                selection1.Add(hybridShapeIntersection)
            else:
                continue
        selection1.Copy()
        self.paste_to_jd(1)

    def create_jd_vectors(self, jd_number):

        part1 = self.select_carm_to_paste_data()
        hybridBodies1 = part1.HybridBodies
        hybridBody1 = hybridBodies1.Item("Joint Definitions")
        hybridBodies2 = hybridBody1.HybridBodies
        hybridBody2 = hybridBodies2.Item("Joint Definition" + ' 0' + str(jd_number))
        hybridShapes1 = hybridBody2.HybridShapes
        hybridShapePointCenter1 = hybridShapes1.Item(1)
        reference1 = part1.CreateReferenceFromObject(hybridShapePointCenter1)
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridBody3 = hybridBodies1.Item("Construction Geometry (REF)")
        hybridBodies3 = hybridBody3.HybridBodies
        hybridBody4 = hybridBodies3.Item("Misc Construction Geometry")
        hybridShapes2 = hybridBody4.HybridShapes
        hybridShapePlaneOffset1 = hybridShapes2.Item('jd' + str(jd_number) + '_vector_direction')
        reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirection(reference2)
        hybridShapeLinePtDir1 = hybridShapeFactory1.AddNewLinePtDir(reference1, hybridShapeDirection1, 0.000000, 25.400000, True)
        hybridBody2.AppendHybridShape(hybridShapeLinePtDir1)
        hybridShapeLinePtDir1.Name = 'FIDV_0' + str(jd_number)
        part1.Update()

    def copy_jd1_fcm10f5cps05wh_and_paste(self, size, type_of_geometry='points'):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.select_current_product()
        selection1.Add(product1)
        if type_of_geometry == 'points':
            selection1.Search(str('(Name = ' + size + '*FAIRING*LIGHT*-Name = *CENTERLINE*), sel'))
            selection1.Copy()
            self.paste_to_jd(1)
        else:
            #selection1.Search(str('(Name = ' + size + '*FAIRING*LIGHT*1_CENTERLINE*), sel'))
            selection1.Search(str('(Name = ' + size + '*FAIRING*LIGHT*-Name = *CENTERLINE*), sel'))
            first_elem = selection1.Item2(1)
            first_point = first_elem.Value
            print first_point.Name
            return first_point


    def activate_view(self, jd_number):

        carm_part = self.select_carm_to_paste_data()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        TPSViews = ann_set1.TPSViews
        view_to_activate = TPSViews.Item(int(jd_number) + 3)
        annotationFactory1 = ann_set1.AnnotationFactory
        ann_set1.ActiveView = view_to_activate
        #annotationFactory1.ActivateTPSView(ann_set1, view_to_activate)

    def add_jd_annotation(self, jd_number, sta_value_pairs, size, side):
        """Adds JOINT DEFINITION XX annotation"""

        annot_text = 'JOINT DEFINITION ' + jd_number
        carm_part = self.access_carm()
        self.activate_view(jd_number)
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        userSurfaces1 = carm_part.UserSurfaces
        geosets = carm_part.HybridBodies
        geoset1 = geosets.Item('Joint Definitions')
        geosets1 = geoset1.HybridBodies
        geoset2 = geosets1.Item('Joint Definition ' + jd_number)
        points = geoset2.HybridShapes
        wb = str(self.workbench_id())
        if wb != 'PrtCfg':
            self.swich_to_part_design()
        reference1 = carm_part.CreateReferenceFromObject(points.Item(1))
        userSurface1 = userSurfaces1.Generate(reference1)

        for point in xrange(2, points.Count):
            reference2 = carm_part.CreateReferenceFromObject(points.Item(point))
            print reference2.name
            userSurface1.AddReference(reference2)

        annotationFactory1 = ann_set1.AnnotationFactory
        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        if jd_number == '01':
            y = Inch_to_mm(12)
            z = 0
            if side == 'LH':
                k = 1
            else:
                k = -1
            #addition = k*(Inch_to_mm(float(size)/2))
            addition = k*(Inch_to_mm(12.0))
        else:
            y = Inch_to_mm(19)
            z = 0
            if side == 'LH':
                k = -1
            else:
                k = 1
            #addition = k*(Inch_to_mm(float(size)/2))
            addition = k*(Inch_to_mm(12.0))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 16)

        self.rename_part_body()

        self.hide_last_annotation()
        carm_part.Update()

        #self.activate_top_prod()



    def manage_annotations_visibility(self):

        part1 = self.select_carm_to_paste_data()
        part1_annotation_set1 = part1.AnnotationSets
        annotation_set1 = part1_annotation_set1.Item(1)
        captures1 = annotation_set1.Captures
        engineering_def_release = captures1.Item(2)
        engineering_def_release_annotations1 = engineering_def_release.Annotations
        engineering_def_release_annotation1 = engineering_def_release_annotations1.Item(1)
        print engineering_def_release_annotation1.Name
        engineering_def_release_annotation1.visProperties.SetShow(0)
        part1.Update()

    def hide_last_annotation(self):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        carm_part = self.select_carm_to_paste_data()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        anns = ann_set1.Annotations
        sta_annotation = anns.Item(anns.Count)
        selection1.Add(sta_annotation)
        selection1.visProperties.SetShow(1)
        carm_part.Update()

    def swich_to_part_design(self):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        part1 = self.select_carm_to_paste_data()
        selection1.Add(part1)
        self.catia.StartWorkbench("PrtCfg")
        print 'part design'

    def unused_activate_top_prod(self):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        top_prod_doc = self.documents.Item(1)

        #top_prod = top_prod_doc.Product

        top_prod1 = top_prod_doc.Parent
        print top_prod1.name
        selection1.Add(top_prod1)
        self.catia.StartWorkbench("Assembly")

        #top_prod2 = top_prod1.Parent
        #print top_prod2.name

    def activate_top_prod(self):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.productDocument1.Product
        print product1.Name
        print product1.PartNumber
        selection1.Add(product1)
        target = selection1.Item(1)
        print target.Name
        self.catia.StartCommand('FrmActivate')
        #wsc = win32com.client.Dispatch("WScript.Shell")
        #wsc.AppActivate("CATIA V5")
        #wsc.SendKeys("c:FrmActivate")
        #wsc.SendKeys("{ENTER}")
        wb = self.workbench_id()
        if wb != 'Assembly':
            self.activate_top_prod()


    def workbench_id(self):

        wb_name = self.catia.GetWorkbenchId()
        print str(wb_name)
        return str(wb_name)


    def hide_unhide_captures(self, hide_or_unhide, capture_number):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        carm_part = self.select_carm_to_paste_data()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        captures1 = ann_set1.Captures
        capture1 = captures1.Item(capture_number)
        selection1.Add(capture1)
        if hide_or_unhide == 'hide':
            selection1.visProperties.SetShow(1)
        else:
            selection1.visProperties.SetShow(0)
        carm_part.Update()
        print 'edr_unhidden'

    def hide_unhide_annotations(self, hide_or_unhide, annotation_number):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        carm_part = self.select_carm_to_paste_data()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        annotations1 = ann_set1.Annotations
        annotation1 = annotations1.Item(annotation_number)
        selection1.Add(annotation1)
        if hide_or_unhide == 'hide':
            selection1.visProperties.SetShow(1)
        else:
            selection1.visProperties.SetShow(0)
        carm_part.Update()
        print 'annotatios visibility managed'



    def map_camera_names(self):
        """Maps camera names to numbers in order"""

        cam_dict = {}
        #carm_part = self.select_carm_to_paste_data()
        documents1 = self.catia.Documents
        partDocument1 = documents1.Item('CA' + self.carm_part_number + '.CATPart')
        cameras = partDocument1.Cameras
        for i in xrange(1, cameras.Count+1):
            camera = cameras.Item(i)
            cam_dict[str(camera.name)] = i
        print cam_dict
        return cam_dict

    def shift_camera(self, sta_value_pairs, size=None):

        #carm_part = self.select_carm_to_paste_data()
        cam_dict = self.map_camera_names()
        documents1 = self.catia.Documents
        partDocument1 = documents1.Item('CA' + self.carm_part_number + '.CATPart')
        cameras = partDocument1.Cameras

        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        coefficient = coord_to_move - (Inch_to_mm(717.0 - float(size)/2))
        print coefficient
        if self.side == 'LH':
            all_annotations_origin = [20118.056641+coefficient, 304.427582, 9348.172852]
            engineering_definition_release_origin = [20120.964844+coefficient, 283.42627, 9091.376953]
            reference_geometry_origin = [18193.632824+coefficient, 329.830675, 9625.886719]
            upper_plenum_fasteners_origin = [18212.539063+coefficient, -5217.299805, 8520.522461]
            upper_plenum_spud_fasteners_origin = [18247.126963+coefficient, 279.302231, 9507.394531]
            all_annotations_sight_direction = [-0.57735, -0.57735, -0.57735]
            all_annotations_up_direction = [-0.4082489, -0.408248, 0.816497]
            engineering_definition_sight_direction = [-0.57735, -0.57735, -0.57735]
            engineering_definition_up_direction = [-0.408248, -0.408248, 0.816497]
            reference_geometry_sight_direction = [0, -0.677411, -0.735605]
            reference_geometry_up_direction = [0, -0.735605, 0.67741]
            upper_plenum_fasteners_sight_direction = [0, 0.941916, -0.335417]
            upper_plenum_fasteners_up_direction = [0, 0.335459, 0.942055]
            upper_plenum_spud_fasteners_sight_direction = [0, -0.67879, -0.734332]
            upper_plenum_spud_fasteners_up_direction = [0, -0.734331, 0.67879]

        else:
            all_annotations_origin = [21650.115234+coefficient, -1780.692627, 9774.560547]
            engineering_definition_release_origin = [21262.044922+coefficient, -1382.012573, 9473.382813]
            reference_geometry_origin = [18193.632824+coefficient, -258.215088, 9591.541016]
            upper_plenum_fasteners_origin = [18504.525391+coefficient, 4084.549561, 8021.676270]
            upper_plenum_spud_fasteners_origin = [18247.126963+coefficient, -472.800446, 9663.113281]
            all_annotations_sight_direction = [-0.64431, 0.635349, -0.425672]
            all_annotations_up_direction = [-0.228644, 0.371113, 0.899998]
            engineering_definition_sight_direction = [-0.63536, 0.63216, -0.443499]
            engineering_definition_up_direction = [-0.294015, 0.333029, 0.895905]
            reference_geometry_sight_direction = [0, 0.677411, -0.735605]
            reference_geometry_up_direction = [0, 0.735605, 0.677411]
            upper_plenum_fasteners_sight_direction = [0, -0.946231, -0.323493]
            upper_plenum_fasteners_up_direction = [0, -0.323493, 0.946231]
            upper_plenum_spud_fasteners_sight_direction = [0, 0.677411, -0.735605]
            upper_plenum_spud_fasteners_up_direction = [0, 0.735605, 0.677411]

        view_origins = [all_annotations_origin, engineering_definition_release_origin, reference_geometry_origin, upper_plenum_fasteners_origin, upper_plenum_spud_fasteners_origin]
        viewpoints = ['All Annotations', 'Engineering Definition Release', 'Reference Geometry', 'Upper Plenum Fasteners', 'Upper Plenum Spud Fasteners']
        sight_directions = [all_annotations_sight_direction, engineering_definition_sight_direction, reference_geometry_sight_direction, upper_plenum_fasteners_sight_direction, upper_plenum_spud_fasteners_sight_direction]
        up_directions = [all_annotations_up_direction, engineering_definition_up_direction, reference_geometry_up_direction, upper_plenum_fasteners_up_direction, upper_plenum_spud_fasteners_up_direction]
        dict_cameras = dict(zip(viewpoints, view_origins))
        dict_sight_directions = dict(zip(viewpoints, sight_directions))
        dict_up_directions = dict(zip(viewpoints, up_directions))
        print dict_cameras
        for view in viewpoints:
            viewpoint = cameras.Item(cam_dict[view])
            print viewpoint.Name
        #sight_direction = [1,0,0]
        #PutUpDirection = [1,1,1]
        #print viewpoint.Viewpoint3D.FocusDistance
        #print viewpoint.Viewpoint3D.Zoom
            vpd = viewpoint.Viewpoint3D
            vpd.PutOrigin(dict_cameras[view])
            vpd.PutSightDirection(dict_sight_directions[view])
            vpd.PutUpDirection(dict_up_directions[view])


        #viewpoint.Viewpoint3D.PutSightDirection(sight_direction)
        #viewpoint.Viewpoint3D.PutUpDirection(sight_direction)

    def paste_to_jd(self, jd_number):
        """Pastes special copied data to JD"""

        selection2 = self.productDocument1.Selection
        selection2.Clear()
        part2 = self.select_carm_to_paste_data()
        hybridBodies1 = part2.HybridBodies
        hybridBody1 = hybridBodies1.Item("Joint Definitions")
        hybridBodies2 = hybridBody1.HybridBodies
        hybridBody2 = hybridBodies2.Item("Joint Definition" + ' 0' + str(jd_number))
        selection2.Add(hybridBody2)
        selection2.PasteSpecial('CATPrtResultWithOutLink')
        selection2.Clear()
        part2.Update()

    def access_captures(self, number):

        selection3 = self.productDocument1.Selection
        selection3.Clear()
        carm_part = self.select_carm_to_paste_data()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        captures1 = ann_set1.Captures
        capture1 = captures1.Item(number)
        print capture1.name
        capture1.ManageHideShowBody = True
        capture1.DisplayCapture()
        #pointers = capture1.Annotations
        #print pointers.name
        #for pointer in pointers:
            #leader = pointers.Item(pointers.Count)
            #print leader.name
        #anns = ann_set1.Annotations
        #jd1_annotation = anns.Item(4)
        #selection3.Add(jd1_annotation)
        #selection3.visProperties.SetShow(0)


    def unused_find_parent(self, child):

        ActiveProductDocument = self.catia.ActiveDocument
        documents1 = self.catia.Documents
        #for i in xrange(1, documents1.Count+1):
            #print documents1.Item(i).name
        Product = self.productDocument1.Product
        Products = Product.Products
        IC = Products.Item(Products.count)
       # print IC.PartNumber
        IC_ref = IC.ReferenceProduct
        my_part = IC_ref.Products.Item(2)
        print my_part.name
        print my_part.PartNumber
        my_parent = my_part.Parent
        print my_parent.name
        Selection = ActiveProductDocument.Selection

    def unused_jd1_bacs12fa3k3(self, copy_from, paste_to):

        documents1 = self.catia.Documents
        partDocument1 = documents1.Item(paste_to)
        part1 = partDocument1.Part
        hybridShapeFactory1 = part1.HybridShapeFactory
        bodies1 = part1.Bodies
        body1 = bodies1.Item("BACS12FA3K3 REF")
        shapes1 = body1.Shapes
        solid1 = shapes1.Item("Solid.22")
        hybridBodies1 = part1.HybridBodies
        hybridBody1 = hybridBodies1.Item("Joint Definitions")
        hybridBodies2 = hybridBody1.HybridBodies
        hybridBody2 = hybridBodies2.Item("Joint Definition 01")
        reference2 = part1.CreateReferenceFromBRepName("REdge:(Edge:(Face:(Brp:(Solid.22;%18);None:();Cf11:());Face:(Brp:(Solid.22;%12);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", solid1)
        hybridShapePointCenter1 = hybridShapeFactory1.AddNewPointCenter(reference2)
        hybridBody2.AppendHybridShape(hybridShapePointCenter1)
        part1.InWorkObject = hybridShapePointCenter1
        part1.Update()


class ProductECS(object):

    def __init__(self):

        self.catia = win32com.client.Dispatch('catia.application')
        self.oFileSys = self.catia.FileSystem
        self.productDocument1 = self.catia.ActiveDocument
        self.documents = self.catia.Documents

    def irms_mapper(self):
        """Builds dict containing instance IDs and indexes"""
        irms_mapping = {}
        product1 = self.productDocument1.Product
        collection_irms = product1.Products

        for prod in xrange(1, collection_irms.Count+1):
            product_inwork = collection_irms.Item(prod)
            irms_mapping[str(product_inwork.name)] = prod
        return irms_mapping

    def unused_text_finder(self, string_to_find, irms):
        """Finds 'string_to_find' in dict built by irms_mapper() and adds them to irms_mapping list"""
        omfs_listing = []
        omfs_listing_keys = irms.keys()
        print omfs_listing_keys
        for text in omfs_listing_keys:
            if string_to_find in text:
                omfs_listing.append(irms[text])
            else:
                continue
        omfs_listing.sort()
        return omfs_listing

    def form_dict(self):
        """Forms dictionary of template names and fairing sizes"""

        numbers = ['Twenty_four_fairing_arch_RH', 'Twenty_four_fairing_arch_LH', 'Twenty_four_fairing.CATPart', 'Twelve', 'Thirty_six_fairing_arch_RH', 'Thirty_six_fairing_arch_LH', 'Thirty_six_fairing', 'Thirty_arch_RH', 'Thirty_arch_LH', 'Thirty', 'Sixty_arch_RH', 'Sixty_arch_LH', 'Sixty', 'Seventy_two', 'Fourty_two_fairing_arch_RH', 'Fourty_two_fairing_arch_LH', 'Fourty_two_fairing', 'Fourty_eight_fairing_arch_RH', 'Fourty_eight_fairing_arch_LH', 'Fourty_eight_fairing', 'Fifty_four_arch_RH', 'Fifty_four_arch_LH', 'Fifty_four', 'Eighteen_arch_RH', 'Eighteen_arch_LH', 'Eighteen']
        words = ['24RH', '24LH', '24', '12', '36RH', '36LH', '36', '30RH', '30LH', '30', '60RH', '60LH', '60', '72', '42RH', '42LH', '42', '48RH', '48LH', '48', '54RH', '54LH', '54', '18RH', '18LH', '18']
        files_dict = dict(zip(words, numbers))
        return files_dict

    def template_finder(self, key):
        """Returns template part number using given key"""
        files_dict = self.form_dict()
        template = files_dict[key]
        extension = '.CATPart'
        return template + extension

class CARM_UPR(CARM):

#def __init__(self, side, *args, **kwargs):
        #super(CARM_UPR, self).__init__(*args, **kwargs)

    def __init__(self, carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name, *args):
        super(CARM_UPR, self).__init__(carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name)
        self.side = side
        self.extention = '\seed_carm_upr_' + side + '.CATPart'
        for i in args:
            if len(i) != 0:
                self.first_elem_in_irm = i[0]
                self.irm_length = len(i)
                self.first_elem_in_irm_size = self.first_elem_in_irm[:2]

    def select_first_elem_in_irm_product(self):
        #ICM_1.ApplyWorkMode(2)
        product1 = self.productDocument1.Product
        products1 = product1.Products
        print 'FIRST ELEMENT IN IRM'
        print self.first_elem_in_irm
        for prod in xrange(1, 5):
            product_to_replace = products1.Item(prod)
            products_to_replace = product_to_replace.Products

            for det in xrange(1, products_to_replace.Count+1):
                product_act_to_replace_nonc = products_to_replace.Item(det)
                if self.first_elem_in_irm in str(product_act_to_replace_nonc.Name):
                    return product_act_to_replace_nonc
                else:
                    continue

        for prod in xrange(5, products1.Count+1):
            product_to_replace = products1.Item(prod)
            if self.first_elem_in_irm in str(product_to_replace.Name):
                return product_to_replace
            else:
                continue


    def select_carm_to_paste_data(self):
        """Returns part of the CARM through the reference product"""

        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        Product2 = to_p.ReferenceProduct
        Product2Products = Product2.Products
        product_forpaste = Product2Products.Item(3)
        print product_forpaste.name
        Part3 = product_forpaste.ReferenceProduct
        PartDocument3 = Part3.Parent
        print PartDocument3.name
        geom_elem3 = PartDocument3.Part
        return geom_elem3

    def change_inst_id(self):

        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        Product2 = to_p.ReferenceProduct
        Product2Products = Product2.Products
        product_forpaste = Product2Products.Item(3)
        product_forpaste.Name = self.instance_id
        print product_forpaste.Name

    def add_ref_annotation(self, sta_value_pairs, size, side):
        """Adds REF annotation"""

        annot_text = str(size) + 'IN OUTBD FRNG SUPPORT REF'
        carm_part = self.access_carm()
        self.activate_ref_view()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        userSurfaces1 = carm_part.UserSurfaces
        geosets = carm_part.HybridBodies
        geoset1 = geosets.Item('Construction Geometry (REF)')
        geosets1 = geoset1.HybridBodies
        geoset2 = geosets1.Item('Misc Construction Geometry')
        hybridShapeFactory1 = carm_part.HybridShapeFactory
        coord_to_move_ref_point = sta_value_pairs[self.copy_from_product - 5] + (Inch_to_mm(float(size))*0.7)
        if side == 'LH':
            coefnt = -1
        else:
            coefnt = 1
        hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(coord_to_move_ref_point, coefnt*1878.360480, 7493.000000)
        geoset2.AppendHybridShape(hybridShapePointCoord1)
        carm_part.Update()
        points = geoset2.HybridShapes
        wb = str(self.workbench_id())
        if wb != 'PrtCfg':
            self.swich_to_part_design()
        reference1 = carm_part.CreateReferenceFromObject(points.Item(points.Count))
        r = points.Item(points.Count)
        print r.Name
        userSurface1 = userSurfaces1.Generate(reference1)
        annotationFactory1 = ann_set1.AnnotationFactory
        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        y = Inch_to_mm(30)
        z = 0
        if side == 'LH':
            k = -1
        else:
            k = 1
        addition = k*(Inch_to_mm(float(size)/2))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 24)
        self.rename_part_body()
        self.hide_last_annotation()
        carm_part.Update()

    def add_sta_annotation(self, sta_value_pairs, sta_values_fake, size, side):
        """Adds REF annotation"""

        sta = sta_values_fake[self.copy_from_product - 5]
        annot_text = 'STA ' + sta + '\n  REF'
        carm_part = self.access_carm()
        self.activate_ref_view()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        userSurfaces1 = carm_part.UserSurfaces
        geosets = carm_part.HybridBodies
        geoset1 = geosets.Item('Construction Geometry (REF)')
        geosets1 = geoset1.HybridBodies
        geoset2 = geosets1.Item('Misc Construction Geometry')
        points = geoset2.HybridShapes
        hybridShapeFactory1 = carm_part.HybridShapeFactory
        coord_to_move_ref_point = sta_value_pairs[self.copy_from_product - 5] + Inch_to_mm(0.25)
        if side == 'LH':
            coefnt = -1
        else:
            coefnt = 1
        hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(coord_to_move_ref_point, coefnt*1878.360480, 7493.000000)
        geoset2.AppendHybridShape(hybridShapePointCoord1)
        carm_part.Update()
        points = geoset2.HybridShapes
        wb = str(self.workbench_id())
        if wb != 'PrtCfg':
            self.swich_to_part_design()
        reference1 = carm_part.CreateReferenceFromObject(points.Item(points.Count))
        userSurface1 = userSurfaces1.Generate(reference1)
        annotationFactory1 = ann_set1.AnnotationFactory
        coord_to_move = sta_value_pairs[self.copy_from_product - 5] - Inch_to_mm(2)
        y = Inch_to_mm(20)
        z = 0
        if side == 'LH':
            k = -1
        else:
            k = 1
        #addition = k*(Inch_to_mm(float(size)/2))*0
        addition = k*(Inch_to_mm(2.15))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 24)
        print ann1text_2d.AnchorPosition
        ann1text_2d.AnchorPosition = 6
        print ann1text_2d.AnchorPosition
        ann1text_2d.FrameType = 3
        self.rename_part_body()
        self.hide_last_annotation()
        carm_part.Update()


    def activate_ref_view(self, jd_number=2):

        carm_part = self.select_carm_to_paste_data()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        TPSViews = ann_set1.TPSViews
        view_to_activate = TPSViews.Item(int(jd_number) + 3)
        annotationFactory1 = ann_set1.AnnotationFactory
        ann_set1.ActiveView = view_to_activate
        #annotationFactory1.ActivateTPSView(ann_set1, view_to_activate)

    def copy_jd1_fcm10f5cps05wh_and_paste(self, size, type_of_geometry='points'):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        if type_of_geometry == 'points':
            product1 = self.select_current_product()
            selection1.Add(product1)
            selection1.Search(str('(Name = ' + size + '*BIN*LIGHT*-Name = *CENTERLINE*), sel'))
            selection1.Copy()
            self.paste_to_jd(1)
        else:
            #selection1.Search(str('(Name = ' + size + '*BIN*LIGHT*1_CENTERLINE*), sel'))
            product1 = self.select_first_elem_in_irm_product()
            selection1.Add(product1)
            selection1.Search(str('(Name = ' + size + '*BIN*LIGHT*-Name = *CENTERLINE*), sel'))
            first_elem = selection1.Item2(1)
            first_point = first_elem.Value
            print first_point.Name
            return first_point


    def copy_jd2_bacs12fa3k3_and_paste(self, size, arch, type_of_geometry='points'):

        selection1 = self.productDocument1.Selection
        selection1.Clear()

        if type_of_geometry == 'points':
            product1 = self.select_current_product()
            selection1.Add(product1)
            if arch:
                selection1.Search(str('(Name = ' + size + '*ARCH*PLENUM*UPR*-Name = *CENTERLINE*), sel'))

            else:
                selection1.Search(str('(Name = ' + size + '*BIN*PLENUM*UPR*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
            selection1.Copy()
            self.paste_to_jd(2)

        else:
            product1 = self.select_first_elem_in_irm_product()
            selection1.Add(product1)
            if arch:
                selection1.Search(str('(Name = ' + size + '*ARCH*PLENUM*UPR*-Name = *CENTERLINE*), sel'))

            else:
                selection1.Search(str('(Name = ' + size + '*BIN*PLENUM*UPR*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
            first_elem = selection1.Item2(1)
            first_point = first_elem.Value
            print first_point.Name
            return first_point



    def shift_camera(self, sta_value_pairs, size=None):

        #carm_part = self.select_carm_to_paste_data()
        cam_dict = self.map_camera_names()
        documents1 = self.catia.Documents
        partDocument1 = documents1.Item('CA' + self.carm_part_number + '.CATPart')
        cameras = partDocument1.Cameras

        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        coefficient = coord_to_move - (Inch_to_mm(717.0 - float(size)/2))
        print coefficient
        if self.side == 'LH':
            all_annotations_origin = [26064.580078+coefficient, 2511.071533, 12036.463867]
            engineering_definition_release_origin = [25083.328125+coefficient, 2428.004395, 11772.291016]
            reference_geometry_origin = [19842.460938+coefficient, 2119.930908, 11485.105469]
            upper_plenum_fasteners_origin = [19780.910156+coefficient, -8050.014648, 9395.449219]
            upper_plenum_spud_fasteners_origin = [20067.941406+coefficient, 2222.089111, 11494.928711]
            all_annotations_sight_direction = [-0.57735, -0.57735, -0.57735]
            all_annotations_up_direction = [-0.4082489, -0.408248, 0.816497]
            engineering_definition_sight_direction = [-0.57735, -0.57735, -0.57735]
            engineering_definition_up_direction = [-0.408248, -0.408248, 0.816497]
            reference_geometry_sight_direction = [0, -0.677411, -0.735605]
            reference_geometry_up_direction = [0, -0.735605, 0.67741]
            upper_plenum_fasteners_sight_direction = [0, 0.941916, -0.335417]
            upper_plenum_fasteners_up_direction = [0, 0.335459, 0.942055]
            upper_plenum_spud_fasteners_sight_direction = [0, -0.67879, -0.734332]
            upper_plenum_spud_fasteners_up_direction = [0, -0.734331, 0.67879]

        else:
            all_annotations_origin = [26915.755859+coefficient, -3865.631592, 10862.578125]
            engineering_definition_release_origin = [28407.517578+coefficient, -5007.974121, 11549.362305]
            reference_geometry_origin = [19574.328125+coefficient, -2151.801514, 11369.263672]
            upper_plenum_fasteners_origin = [19731.783203+coefficient, 7120.067383, 6091.5625]
            upper_plenum_spud_fasteners_origin = [19156.607422+coefficient, -1839.441162, 10945.449219]
            all_annotations_sight_direction = [-0.64431, 0.635349, -0.425672]
            all_annotations_up_direction = [-0.228644, 0.371113, 0.899998]
            engineering_definition_sight_direction = [-0.64431, 0.635349, -0.425672]
            engineering_definition_up_direction = [-0.228644, 0.371113, 0.899998]
            reference_geometry_sight_direction = [0, 0.677411, -0.735605]
            reference_geometry_up_direction = [0, 0.735605, 0.677411]
            upper_plenum_fasteners_sight_direction = [0, -0.981109, 0.193369]
            upper_plenum_fasteners_up_direction = [0, 0.19337, 0.981126]
            upper_plenum_spud_fasteners_sight_direction = [0, 0.677411, -0.735605]
            upper_plenum_spud_fasteners_up_direction = [0, 0.735605, 0.677411]

        view_origins = [all_annotations_origin, engineering_definition_release_origin, reference_geometry_origin, upper_plenum_fasteners_origin, upper_plenum_spud_fasteners_origin]
        viewpoints = ['All Annotations', 'Engineering Definition Release', 'Reference Geometry', 'Upper Plenum Fasteners', 'Upper Plenum Spud Fasteners']
        sight_directions = [all_annotations_sight_direction, engineering_definition_sight_direction, reference_geometry_sight_direction, upper_plenum_fasteners_sight_direction, upper_plenum_spud_fasteners_sight_direction]
        up_directions = [all_annotations_up_direction, engineering_definition_up_direction, reference_geometry_up_direction, upper_plenum_fasteners_up_direction, upper_plenum_spud_fasteners_up_direction]
        dict_cameras = dict(zip(viewpoints, view_origins))
        dict_sight_directions = dict(zip(viewpoints, sight_directions))
        dict_up_directions = dict(zip(viewpoints, up_directions))
        print dict_cameras
        for view in viewpoints:
            viewpoint = cameras.Item(cam_dict[view])
            print viewpoint.Name
        #sight_direction = [1,0,0]
        #PutUpDirection = [1,1,1]
        #print viewpoint.Viewpoint3D.FocusDistance
        #print viewpoint.Viewpoint3D.Zoom
            vpd = viewpoint.Viewpoint3D
            vpd.PutOrigin(dict_cameras[view])
            vpd.PutSightDirection(dict_sight_directions[view])
            vpd.PutUpDirection(dict_up_directions[view])


        #viewpoint.Viewpoint3D.PutSightDirection(sight_direction)
        #viewpoint.Viewpoint3D.PutUpDirection(sight_direction)

    def set_parameters(self, sta_value_pairs, size):

        carm_part = self.select_carm_to_paste_data()
        parameters1 = carm_part.Parameters
        ref_param = parameters1.Item('ref_connector_X')
        sta_param = parameters1.Item('sta_connector_X')
        #direct_param = parameters1.Item('view_direction_connector_X')
        print ref_param.Value
        print sta_param.Value
        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        print coord_to_move
        ref_param.Value = coord_to_move + (Inch_to_mm(float(size))) - (Inch_to_mm(float(size))*0.3)
        sta_param.Value = coord_to_move + Inch_to_mm(0.25)
        #direct_param.Value = coord_to_move + (Inch_to_mm(float(size)/2)) + Inch_to_mm(7.0)
        print ref_param.Value
        print sta_param.Value

    def add_jd_annotation(self, jd_number, sta_value_pairs, size, side, arch):
        """Adds JOINT DEFINITION XX annotation"""

        annot_text = 'JOINT DEFINITION ' + jd_number
        carm_part = self.access_carm()
        self.activate_view(jd_number)
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        userSurfaces1 = carm_part.UserSurfaces
        geosets = carm_part.HybridBodies
        geoset1 = geosets.Item('Joint Definitions')
        geosets1 = geoset1.HybridBodies
        geoset2 = geosets1.Item('Joint Definition ' + jd_number)
        points = geoset2.HybridShapes
        if '1' in str(jd_number):
            JD_point = self.copy_jd1_fcm10f5cps05wh_and_paste(self.first_elem_in_irm_size, 'find_point')
        else:
            JD_point = self.copy_jd2_bacs12fa3k3_and_paste(self.first_elem_in_irm_size, False, 'find_point')
        #JD_point = points.Item(1)
        #hybridShapeFactory1 = carm_part.HybridShapeFactory
        #fake_JD_point = hybridShapeFactory1.AddNewPointCoordWithReference(0, 0, 0, JD_point)
        JD_point_coord_X = JD_point.X
        JD_point_X = JD_point_coord_X.Value
        print JD_point_X
        wb = str(self.workbench_id())
        if wb != 'PrtCfg':
            self.swich_to_part_design()
        reference1 = carm_part.CreateReferenceFromObject(points.Item(1))
        userSurface1 = userSurfaces1.Generate(reference1)

        for point in xrange(2, points.Count):
            reference2 = carm_part.CreateReferenceFromObject(points.Item(point))
            print reference2.name
            userSurface1.AddReference(reference2)

        annotationFactory1 = ann_set1.AnnotationFactory
        #coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        #coord_to_move = sta_value_pairs[self.copy_from_product - 5 - (self.irm_length - 1)]
        coord_to_move = sta_value_pairs[self.copy_from_product - 5 - (self.irm_length - 1)] + JD_point_X
        print coord_to_move
        if jd_number == '01':
            y = Inch_to_mm(12)
            z = 0
            if side == 'LH':
                k = 1
            else:
                k = -1
            #addition = k*(Inch_to_mm(float(size)/2))
            addition = k*Inch_to_mm(12.0)
        else:
            y = Inch_to_mm(19)
            z = 0
            if side == 'LH':
                k = -1
            else:
                k = 1
            #addition = k*(Inch_to_mm(float(size)/2))
            addition = k*Inch_to_mm(12.0)
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 24)

        self.rename_part_body()

        self.hide_last_annotation()
        carm_part.Update()

        #self.activate_top_prod()


class CARM_UPR_NONC(CARM):

#def __init__(self, side, *args, **kwargs):
        #super(CARM_UPR, self).__init__(*args, **kwargs)

    def __init__(self, carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name):
        super(CARM_UPR_NONC, self).__init__(carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name)
        self.side = side
        self.extention = '\seed_carm_upr_' + side + '.CATPart'

    def select_carm_to_paste_data(self):
        """Returns part of the CARM through the reference product"""

        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        Product2 = to_p.ReferenceProduct
        Product2Products = Product2.Products
        product_forpaste = Product2Products.Item(3)
        print product_forpaste.name
        Part3 = product_forpaste.ReferenceProduct
        PartDocument3 = Part3.Parent
        print PartDocument3.name
        geom_elem3 = PartDocument3.Part
        return geom_elem3

    def change_inst_id(self):

        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        Product2 = to_p.ReferenceProduct
        Product2Products = Product2.Products
        product_forpaste = Product2Products.Item(3)
        product_forpaste.Name = self.instance_id
        print product_forpaste.Name

    def add_ref_annotation(self, sta_value_pairs, size, side):
        """Adds REF annotation"""

        annot_text = str(size) + 'IN OUTBD FRNG SUPPORT REF'
        carm_part = self.access_carm()
        self.activate_ref_view()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        userSurfaces1 = carm_part.UserSurfaces
        geosets = carm_part.HybridBodies
        geoset1 = geosets.Item('Construction Geometry (REF)')
        geosets1 = geoset1.HybridBodies
        geoset2 = geosets1.Item('Misc Construction Geometry')
        hybridShapeFactory1 = carm_part.HybridShapeFactory
        coord_to_move_ref_point = sta_value_pairs[self.copy_from_product - 5] + (Inch_to_mm(float(size))/2.0)
        if side == 'LH':
            coefnt = -1
        else:
            coefnt = 1
        hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(coord_to_move_ref_point, coefnt*1878.360480, 7493.000000)
        geoset2.AppendHybridShape(hybridShapePointCoord1)
        carm_part.Update()
        points = geoset2.HybridShapes
        wb = str(self.workbench_id())
        if wb != 'PrtCfg':
            self.swich_to_part_design()
        reference1 = carm_part.CreateReferenceFromObject(points.Item(points.Count))
        r = points.Item(points.Count)
        print r.Name
        userSurface1 = userSurfaces1.Generate(reference1)
        annotationFactory1 = ann_set1.AnnotationFactory
        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        y = Inch_to_mm(28)
        z = 0
        if side == 'LH':
            k = -1
        else:
            k = 1
        addition = k*(Inch_to_mm(float(size)/2))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 20)
        self.rename_part_body()
        self.hide_last_annotation()
        carm_part.Update()

    def add_sta_annotation(self, sta_value_pairs, sta_values_fake, size, side):
        """Adds REF annotation"""

        sta = sta_values_fake[self.copy_from_product - 5]
        annot_text = 'STA ' + sta + '\nREF'
        carm_part = self.access_carm()
        self.activate_ref_view()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        userSurfaces1 = carm_part.UserSurfaces
        geosets = carm_part.HybridBodies
        geoset1 = geosets.Item('Construction Geometry (REF)')
        geosets1 = geoset1.HybridBodies
        geoset2 = geosets1.Item('Misc Construction Geometry')
        points = geoset2.HybridShapes
        hybridShapeFactory1 = carm_part.HybridShapeFactory
        coord_to_move_ref_point = sta_value_pairs[self.copy_from_product - 5]
        if side == 'LH':
            coefnt = -1
        else:
            coefnt = 1
        hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(coord_to_move_ref_point, coefnt*1878.360480, 7493.000000)
        geoset2.AppendHybridShape(hybridShapePointCoord1)
        carm_part.Update()
        points = geoset2.HybridShapes
        wb = str(self.workbench_id())
        if wb != 'PrtCfg':
            self.swich_to_part_design()
        reference1 = carm_part.CreateReferenceFromObject(points.Item(points.Count))
        userSurface1 = userSurfaces1.Generate(reference1)
        annotationFactory1 = ann_set1.AnnotationFactory
        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        y = Inch_to_mm(22)
        z = 0
        if side == 'LH':
            k = -1
        else:
            k = 1
        #addition = k*(Inch_to_mm(float(size)/2))*0
        addition = k*(Inch_to_mm(2.15))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 20)
        ann1text_2d.AnchorPosition = 6
        ann1text_2d.FrameType = 3
        self.rename_part_body()
        self.hide_last_annotation()
        carm_part.Update()


    def activate_ref_view(self, jd_number=2):

        carm_part = self.select_carm_to_paste_data()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        TPSViews = ann_set1.TPSViews
        view_to_activate = TPSViews.Item(int(jd_number) + 3)
        annotationFactory1 = ann_set1.AnnotationFactory
        ann_set1.ActiveView = view_to_activate
        #annotationFactory1.ActivateTPSView(ann_set1, view_to_activate)

    def copy_jd1_fcm10f5cps05wh_and_paste(self, size, type_of_geometry='points'):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.select_current_product()
        selection1.Add(product1)
        if type_of_geometry == 'points':
            selection1.Search(str('(Name = ' + size + '*BIN*LIGHT*-Name = *CENTERLINE*), sel'))
        else:
            selection1.Search(str('(Name = ' + size + '*BIN*LIGHT*1_CENTERLINE*), sel'))
        selection1.Copy()
        self.paste_to_jd(1)

    def copy_jd2_bacs12fa3k3_and_paste(self, size, arch, type_of_geometry='points'):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.select_current_product()
        selection1.Add(product1)
        if type_of_geometry == 'points':
            if arch:
                selection1.Search(str('(Name = ' + size + '*ARCH*PLENUM*UPR*-Name = *CENTERLINE*), sel'))

            else:
                selection1.Search(str('(Name = ' + size + '*BIN*PLENUM*UPR*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))

        selection1.Copy()
        self.paste_to_jd(2)

    def shift_camera(self, sta_value_pairs, size=None):

        #carm_part = self.select_carm_to_paste_data()
        cam_dict = self.map_camera_names()
        documents1 = self.catia.Documents
        partDocument1 = documents1.Item('CA' + self.carm_part_number + '.CATPart')
        cameras = partDocument1.Cameras

        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        coefficient = coord_to_move - (Inch_to_mm(717.0 - float(size)/2))
        print coefficient
        if self.side == 'LH':
            all_annotations_origin = [26064.580078+coefficient, 2511.071533, 12036.463867]
            engineering_definition_release_origin = [25083.328125+coefficient, 2428.004395, 11772.291016]
            reference_geometry_origin = [19842.460938+coefficient, 2119.930908, 11485.105469]
            upper_plenum_fasteners_origin = [19780.910156+coefficient, -8050.014648, 9395.449219]
            upper_plenum_spud_fasteners_origin = [20067.941406+coefficient, 2222.089111, 11494.928711]
            all_annotations_sight_direction = [-0.57735, -0.57735, -0.57735]
            all_annotations_up_direction = [-0.4082489, -0.408248, 0.816497]
            engineering_definition_sight_direction = [-0.57735, -0.57735, -0.57735]
            engineering_definition_up_direction = [-0.408248, -0.408248, 0.816497]
            reference_geometry_sight_direction = [0, -0.677411, -0.735605]
            reference_geometry_up_direction = [0, -0.735605, 0.67741]
            upper_plenum_fasteners_sight_direction = [0, 0.941916, -0.335417]
            upper_plenum_fasteners_up_direction = [0, 0.335459, 0.942055]
            upper_plenum_spud_fasteners_sight_direction = [0, -0.67879, -0.734332]
            upper_plenum_spud_fasteners_up_direction = [0, -0.734331, 0.67879]

        else:
            all_annotations_origin = [26915.755859+coefficient, -3865.631592, 10862.578125]
            engineering_definition_release_origin = [28407.517578+coefficient, -5007.974121, 11549.362305]
            reference_geometry_origin = [19574.328125+coefficient, -2151.801514, 11369.263672]
            upper_plenum_fasteners_origin = [19731.783203+coefficient, 7120.067383, 6091.5625]
            upper_plenum_spud_fasteners_origin = [19156.607422+coefficient, -1839.441162, 10945.449219]
            all_annotations_sight_direction = [-0.64431, 0.635349, -0.425672]
            all_annotations_up_direction = [-0.228644, 0.371113, 0.899998]
            engineering_definition_sight_direction = [-0.64431, 0.635349, -0.425672]
            engineering_definition_up_direction = [-0.228644, 0.371113, 0.899998]
            reference_geometry_sight_direction = [0, 0.677411, -0.735605]
            reference_geometry_up_direction = [0, 0.735605, 0.677411]
            upper_plenum_fasteners_sight_direction = [0, -0.981109, 0.193369]
            upper_plenum_fasteners_up_direction = [0, 0.19337, 0.981126]
            upper_plenum_spud_fasteners_sight_direction = [0, 0.677411, -0.735605]
            upper_plenum_spud_fasteners_up_direction = [0, 0.735605, 0.677411]

        view_origins = [all_annotations_origin, engineering_definition_release_origin, reference_geometry_origin, upper_plenum_fasteners_origin, upper_plenum_spud_fasteners_origin]
        viewpoints = ['All Annotations', 'Engineering Definition Release', 'Reference Geometry', 'Upper Plenum Fasteners', 'Upper Plenum Spud Fasteners']
        sight_directions = [all_annotations_sight_direction, engineering_definition_sight_direction, reference_geometry_sight_direction, upper_plenum_fasteners_sight_direction, upper_plenum_spud_fasteners_sight_direction]
        up_directions = [all_annotations_up_direction, engineering_definition_up_direction, reference_geometry_up_direction, upper_plenum_fasteners_up_direction, upper_plenum_spud_fasteners_up_direction]
        dict_cameras = dict(zip(viewpoints, view_origins))
        dict_sight_directions = dict(zip(viewpoints, sight_directions))
        dict_up_directions = dict(zip(viewpoints, up_directions))
        print dict_cameras
        for view in viewpoints:
            viewpoint = cameras.Item(cam_dict[view])
            print viewpoint.Name
        #sight_direction = [1,0,0]
        #PutUpDirection = [1,1,1]
        #print viewpoint.Viewpoint3D.FocusDistance
        #print viewpoint.Viewpoint3D.Zoom
            vpd = viewpoint.Viewpoint3D
            vpd.PutOrigin(dict_cameras[view])
            vpd.PutSightDirection(dict_sight_directions[view])
            vpd.PutUpDirection(dict_up_directions[view])


        #viewpoint.Viewpoint3D.PutSightDirection(sight_direction)
        #viewpoint.Viewpoint3D.PutUpDirection(sight_direction)


class CARM_LWR(CARM_UPR):

#def __init__(self, side, *args, **kwargs):
        #super(CARM_UPR, self).__init__(*args, **kwargs)

    def __init__(self, carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name, *args):
        super(CARM_LWR, self).__init__(carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name, *args)
        self.side = side
        self.extention = '\seed_carm_lwr_' + side + '.CATPart'

    def copy_jd1_BACS12FA3K20_and_paste(self, size, type_of_geometry='points'):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.select_current_product()
        selection1.Add(product1)
        selection1.Search(str('(Name = ' + size + '*BIN*PLENUM*LWR*BACI12AG3UCM2*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
        if type_of_geometry == 'points':
            selection1.Copy()
            self.paste_to_jd(1)
        else:
            first_elem = selection1.Item2(1)
            first_point = first_elem.Value
            print first_point.Name
            return first_point


    def copy_jd2_bacs12fa3k3_and_paste_1(self, size, type_of_geometry='points'):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.select_current_product()
        selection1.Add(product1)
        selection1.Search(str('(Name = ' + size + '*BIN*PLENUM*LWR*BACI12AK3CM07*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
        if type_of_geometry == 'points':
            selection1.Copy()
            self.paste_to_jd(2)
        else:
            first_elem = selection1.Item2(1)
            first_point = first_elem.Value
            print first_point.Name
            return first_point

    def copy_jd3_BACS12FA3K12_and_paste(self, size, type_of_geometry='points'):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.select_current_product()
        selection1.Add(product1)
        selection1.Search(str('(Name = ' + size + '*NOZZLE*LOWER*BACI12AH5U375*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
        if type_of_geometry == 'points':
            selection1.Copy()
            self.paste_to_jd(3)
        else:
            first_elem = selection1.Item2(1)
            first_point = first_elem.Value
            print first_point.Name
            return first_point

    def copy_jd4_bacs12fa3k3_and_paste_2(self, size, type_of_geometry='points'):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.select_current_product()
        selection1.Add(product1)
        selection1.Search(str('(Name = ' + size + '*OB*BIN*END*FRAME*ECS*NOZZLE*BACI12AK3CM07*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
        if type_of_geometry == 'points':
            selection1.Copy()
            self.paste_to_jd(4)
        else:
            first_elem = selection1.Item2(1)
            first_point = first_elem.Value
            print first_point.Name
            return first_point

    def set_parameters(self, sta_value_pairs, size):

        carm_part = self.select_carm_to_paste_data()
        parameters1 = carm_part.Parameters
        FL2_X_param = parameters1.Item('FL2_X')
        FL3_X_param = parameters1.Item('FL3_X')
        FL4_X_param = parameters1.Item('FL4_X')
        FL5_X_param = parameters1.Item('FL5_X')
        fir_tree_param = parameters1.Item('156-00066')
        bacs_param = parameters1.Item('BACS38K2')
        FL2_X_param_offset = Inch_to_mm(4.0606)
        FL3_X_param_offset = Inch_to_mm(2.1582)
        FL4_X_param_offset = Inch_to_mm(3.0059)
        FL5_X_param_offset = Inch_to_mm(2.4261)
        fir_tree_param_offset = Inch_to_mm(0.4353)
        bacs_param_offset = Inch_to_mm(1.4077)
        anchor_point = self.copy_jd2_bacs12fa3k3_and_paste_1(size, 'find_point')
        anchor_point_coord_X = anchor_point.X
        anchor_point_X = anchor_point_coord_X.Value
        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        print coord_to_move
        FL2_X_param.Value = coord_to_move + anchor_point_X + FL2_X_param_offset
        FL3_X_param.Value = coord_to_move + anchor_point_X + FL3_X_param_offset
        FL4_X_param.Value = coord_to_move + anchor_point_X + FL4_X_param_offset
        FL5_X_param.Value = coord_to_move + anchor_point_X + FL5_X_param_offset
        fir_tree_param.Value = coord_to_move + anchor_point_X + fir_tree_param_offset
        bacs_param.Value = coord_to_move + anchor_point_X + bacs_param_offset

    def shift_camera(self, sta_value_pairs, size=None):

        #carm_part = self.select_carm_to_paste_data()
        cam_dict = self.map_camera_names()
        documents1 = self.catia.Documents
        partDocument1 = documents1.Item('CA' + self.carm_part_number + '.CATPart')
        cameras = partDocument1.Cameras

        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        coefficient = coord_to_move - (Inch_to_mm(717.0 - float(size)/2))
        print coefficient

        if self.side == 'LH':
            engineering_definition_release_origin = [12313.794922 + coefficient, -6951.88916, 9594.439453]
            all_annotations_origin = [10896.336914 + coefficient, -7317.958984, 10955.158203]
            reference_geometry_origin = [15983.475586 + coefficient, 3519.238281, 7186.692383]
            lower_plenum_downer_strap_origin = [19780.910156 + coefficient, -3187.991943, 6838.505859]
            upper_downer_strap_origin = [14542.484375 + coefficient, -2463.961426, 8110.249512]
            lower_plenum_fastener_jd01_origin = [14483.808594 + coefficient, -3314.811768, 7032.354492]
            lower_plenum_fastener_jd02_origin = [14517.706055 + coefficient, -2933.570801, 7694.333984]
            sidewall_nozzle_fastener_jd03_origin = [16085.119141 + coefficient, -1406.840454, 6359.020508]
            sidewall_nozzle_fastener_jd04_origin = [14804.630859 + coefficient, -3773.533447, 6764.081055]

            engineering_definition_sight_direction = [0.567468, 0.725849, -0.388744]
            engineering_definition_up_direction = [0.24319, 0.303315, 0.921335]
            all_annotations_sight_direction = [0.589343, 0.653958, -0.474356]
            all_annotations_up_direction = [0.312794, 0.356658, 0.880315]
            reference_geometry_sight_direction = [0, -1, 0]
            reference_geometry_up_direction = [0, 0, 1]
            lower_plenum_downer_strap_sight_direction = [0, 1, 0]
            lower_plenum_downer_strap_up_direction = [0, 0, 1]
            upper_downer_strap_sight_direction = [0, 0.734878, -0.6782]
            upper_downer_strap_up_direction = [0, 0.6782, 0.734878]
            lower_plenum_fastener_jd01_sight_direction = [0, 0.935714, -0.352759]
            lower_plenum_fastener_jd01_up_direction = [0, 0.352757, 0.935706]
            lower_plenum_fastener_jd02_sight_direction = [0, 0.509744, -0.860024]
            lower_plenum_fastener_jd02_up_direction = [0, 0.86027, 0.509834]
            sidewall_nozzle_fastener_jd03_sight_direction = [0, -0.937411, 0.348038]
            sidewall_nozzle_fastener_jd03_up_direction = [0, 0.348079, 0.937464]
            sidewall_nozzle_fastener_jd04_sight_direction = [0, 1, 0]
            sidewall_nozzle_fastener_jd04_up_direction = [0, 0, 1]

        else:
            engineering_definition_release_origin = [12313.794922 + coefficient, -6951.88916, 9594.439453]
            all_annotations_origin = [10896.336914 + coefficient, -7317.958984, 10955.158203]
            reference_geometry_origin = [15983.475586 + coefficient, 3519.238281, 7186.692383]
            lower_plenum_downer_strap_origin = [19780.910156 + coefficient, -3187.991943, 6838.505859]
            upper_downer_strap_origin = [14542.484375 + coefficient, -2463.961426, 8110.249512]
            lower_plenum_fastener_jd01_origin = [14483.808594 + coefficient, -3314.811768, 7032.354492]
            lower_plenum_fastener_jd02_origin = [14517.706055 + coefficient, -2933.570801, 7694.333984]
            sidewall_nozzle_fastener_jd03_origin = [16085.119141 + coefficient, -1406.840454, 6359.020508]
            sidewall_nozzle_fastener_jd04_origin = [14804.630859 + coefficient, -3773.533447, 6764.081055]

            engineering_definition_sight_direction = [0.567468, 0.725849, -0.388744]
            engineering_definition_up_direction = [0.24319, 0.303315, 0.921335]
            all_annotations_sight_direction = [0.589343, 0.653958, -0.474356]
            all_annotations_up_direction = [0.312794, 0.356658, 0.880315]
            reference_geometry_sight_direction = [0, -1, 0]
            reference_geometry_up_direction = [0, 0, 1]
            lower_plenum_downer_strap_sight_direction = [0, 1, 0]
            lower_plenum_downer_strap_up_direction = [0, 0, 1]
            upper_downer_strap_sight_direction = [0, 0.734878, -0.6782]
            upper_downer_strap_up_direction = [0, 0.6782, 0.734878]
            lower_plenum_fastener_jd01_sight_direction = [0, 0.935714, -0.352759]
            lower_plenum_fastener_jd01_up_direction = [0, 0.352757, 0.935706]
            lower_plenum_fastener_jd02_sight_direction = [0, 0.509744, -0.860024]
            lower_plenum_fastener_jd02_up_direction = [0, 0.86027, 0.509834]
            sidewall_nozzle_fastener_jd03_sight_direction = [0, -0.937411, 0.348038]
            sidewall_nozzle_fastener_jd03_up_direction = [0, 0.348079, 0.937464]
            sidewall_nozzle_fastener_jd04_sight_direction = [0, 1, 0]
            sidewall_nozzle_fastener_jd04_up_direction = [0, 0, 1]

        view_origins = [all_annotations_origin, engineering_definition_release_origin, reference_geometry_origin, lower_plenum_downer_strap_origin, upper_downer_strap_origin, lower_plenum_fastener_jd01_origin, lower_plenum_fastener_jd02_origin, sidewall_nozzle_fastener_jd03_origin, sidewall_nozzle_fastener_jd04_origin]
        viewpoints = ['ALL ANNOTATIONS', 'ENGINEERING DEFINITION RELEASE', 'REFERENCE GEOMETRY', 'LOWER PLENUM DOWNER STRAP', 'UPPER DOWNER STRAP', 'LOWER PLENUM FASTENER JD01', 'LOWER PLENUM FASTENER JD02', 'SIDEWALL NOZZLE FASTENER JD03', 'SIDEWALL NOZZLE FASTENER JD04']
        sight_directions = [all_annotations_sight_direction, engineering_definition_sight_direction, reference_geometry_sight_direction, lower_plenum_downer_strap_sight_direction, upper_downer_strap_sight_direction, lower_plenum_fastener_jd01_sight_direction, lower_plenum_fastener_jd02_sight_direction, sidewall_nozzle_fastener_jd03_sight_direction, sidewall_nozzle_fastener_jd04_sight_direction]
        up_directions = [all_annotations_up_direction, engineering_definition_up_direction, reference_geometry_up_direction, lower_plenum_downer_strap_up_direction, upper_downer_strap_up_direction, lower_plenum_fastener_jd01_up_direction, lower_plenum_fastener_jd02_up_direction, sidewall_nozzle_fastener_jd03_up_direction, sidewall_nozzle_fastener_jd04_up_direction]
        dict_cameras = dict(zip(viewpoints, view_origins))
        dict_sight_directions = dict(zip(viewpoints, sight_directions))
        dict_up_directions = dict(zip(viewpoints, up_directions))
        print dict_cameras
        for view in viewpoints:
            viewpoint = cameras.Item(cam_dict[view])
            print viewpoint.Name
        #sight_direction = [1,0,0]
        #PutUpDirection = [1,1,1]
        #print viewpoint.Viewpoint3D.FocusDistance
        #print viewpoint.Viewpoint3D.Zoom
            vpd = viewpoint.Viewpoint3D
            vpd.PutOrigin(dict_cameras[view])
            vpd.PutSightDirection(dict_sight_directions[view])
            vpd.PutUpDirection(dict_up_directions[view])


        #viewpoint.Viewpoint3D.PutSightDirection(sight_direction)
        #viewpoint.Viewpoint3D.PutUpDirection(sight_direction)

    def activate_view(self, view_number):

        carm_part = self.select_carm_to_paste_data()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        TPSViews = ann_set1.TPSViews
        view_to_activate = TPSViews.Item(view_number)
        ann_set1.ActiveView = view_to_activate

    def select_carm_to_paste_data(self):
        """Returns part of the CARM through the reference product"""

        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        if '24' in to_p.Name:
            order_of_carm_in_tree = 3
        else:
            order_of_carm_in_tree = 4
        Product2 = to_p.ReferenceProduct
        Product2Products = Product2.Products
        product_forpaste = Product2Products.Item(order_of_carm_in_tree)
        print product_forpaste.name
        Part3 = product_forpaste.ReferenceProduct
        PartDocument3 = Part3.Parent
        print PartDocument3.name
        geom_elem3 = PartDocument3.Part
        return geom_elem3

    def set_standard_parts_params(self, jd_number):

        self.update_hole_qty(jd_number)
        carm_part = self.select_carm_to_paste_data()
        parameters1 = carm_part.Parameters
        selection1 = self.productDocument1.Selection
        selection1.Clear()
        hole_qty = self.calculate_jd_points(jd_number)
        if jd_number == 1:
            param1 = parameters1.Item('BACS12FA3K20')
            param1.Value = str(hole_qty) + '|BACS12FA3K20 | SCREW, WASHER HEAD, CROSS RECESS, FULL THREADED, 6AL-4V TITANIUM'
            param2 = parameters1.Item('BACS38K2')
            param2.Value = str(hole_qty/2) + '|BACS38K2 | STRAP, ADJUSTABLE'
            param3 = parameters1.Item('156-00066')
            param3.Value = str(hole_qty) + '|156-00066 | ASSEMBLY TREE MOUNT AND CABLE TIES'
        elif jd_number == 2:
            hole_qty1 = self.calculate_jd_points(4)
            hole_qty += hole_qty1
            param = parameters1.Item('BACS12FA3K3')
            param.Value = str(hole_qty) + '|BACS12FA3K3 | SCREW, WASHER HEAD, CROSS RECESS, FULL THREADED, 6AL-4V TITANIUM'
        elif jd_number == 3:
            param = parameters1.Item('BACS12FA3K12')
            param.Value = str(hole_qty) + '|BACS12FA3K12 | SCREW, WASHER HEAD, CROSS RECESS, FULL THREADED, 6AL-4V TITANIUM'
        elif jd_number == 4:
            pass

    def calculate_jd_points(self, jd_number):

        carm_part = self.select_carm_to_paste_data()
        hole_qty = 0
        selection1 = self.productDocument1.Selection
        selection1.Clear()
        hybridBodies1 = carm_part.HybridBodies
        hybridBody1 = hybridBodies1.Item("Joint Definitions")
        hybridBodies2 = hybridBody1.HybridBodies
        hybridBody2 = hybridBodies2.Item("Joint Definition" + ' 0' + str(jd_number))
        HybridShapes1 = hybridBody2.HybridShapes
        for shape in xrange(HybridShapes1.Count):
            hole_qty += 1
        print hole_qty
        return hole_qty

    def update_hole_qty(self, jd_number):

        carm_part = self.select_carm_to_paste_data()
        parameters1 = carm_part.Parameters
        hole_qty = self.calculate_jd_points(jd_number)
        param_hole_qty = parameters1.Item('Joint Definitions\Joint Definition 0' + str(jd_number) + '\Hole Quantity')
        param_hole_qty.Value = str(hole_qty)



#x1 = CARM('830Z1000-2194', 'INSTL_UPR_AIR_DIST_OMF_CARM', 'LH', 2, 1)
#x2 = CARM('8755675', 2)
#prod = ProductECS()

#group = [prod]

#key = '24RH'

#print prod.template_finder(key)
#sta_value_pairs = [18211.8, 20650.199999999997, 18211.8, 20650.199999999997]
#a = Inch_to_mm(1)
#print a
#for i in group:

#x1.add_carm_as_external_component()
#x1.access_captures(2)
#x1.hide_unhide_annotations('hide', 1)
#x1.hide_unhide_annotations('hide', 2)
#x1.manage_annotations_visibility()
#x1.add_jd_annotation('01', sta_value_pairs, '54', 'LH')
#x1.switch_to_part_design()
#x1.activate_top_prod()

#x1.activate_top_prod()
#x1.change_inst_id()
    #i.rename_carm()
    #i.add_geosets()
    #i.access_annotations()
    #i.get_points()
    #i.set_parameters()
    #i.unused_copy_bodies_and_paste('Part1.CATPart', 'CA03.CATPart')
    #i.unused_copy_ref_surface_and_paste('Part1.CATPart', 'CA03.CATPart')
    #i.unused_copy_jd1_fcm10f5cps05wh_and_paste('Thirty.CATPart', 'seed_carm.CATPart')
    #i.unused_copy_jd2_bacs12fa3k3_and_paste('Thirty.CATPart', 'seed_carm.CATPart')
#x1.map_camera_names()
#x1.shift_camera()
    #i.find_parent('Eighteen_solids301.CATProduct')
    #i.add_jd_annotation(15)
    #d = i.irms_mapper()
    #print d
    #val = i.text_finder('six', d)
    #print val
    #dt = i.form_dict()
    #print dt
    #print dt.keys()
    #print dt['Fifty_four']
