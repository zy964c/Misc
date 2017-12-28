
from APIE_Tool2v10 import inch_to_mm
from APIE_Tool2v10 import sta_value
from APIE_Tool2v10 import var_import
import win32com.client


class CarmOmf(object):

    def __init__(self, carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name):
        self.carm_part_number = carm_part_number
        self.order_of_new_product = order_of_new_product
        self.instance_id = instance_id
        self.copy_from_product = copy_from_product
        self.cfp_name = cfp_name
        self.catia = win32com.client.Dispatch('catia.application')
        self.path = [var_import[0]]
        self.side = side
        self.extention = '\\seed_carm_' + side + '.CATPart'
        self.oFileSys = self.catia.FileSystem
        self.productDocument1 = self.catia.ActiveDocument
        self.documents = self.catia.Documents

    def add_carm_as_external_component(self):
        """Instantiates CARM from external library"""

        product1 = self.productDocument1.Product
        collection_irms = product1.Products
        product_to_insert_carm = collection_irms.Item(self.order_of_new_product)
        children_of_product_to_insert_carm = product_to_insert_carm.Products
        PartDocPath = self.path[0] + self.extention
        print PartDocPath
        PartDocPath1 = self.path[0] + '\CA' + self.carm_part_number + '.CATPart'
        self.oFileSys.CopyFile(PartDocPath, PartDocPath1, True)
        PartDoc = self.catia.Documents.NewFrom(PartDocPath1)
        PartDoc1 = PartDoc.Product
        PartDoc1.PartNumber = 'CA' + self.carm_part_number
        print PartDoc1.Name
        NewComponent = children_of_product_to_insert_carm.AddExternalComponent(PartDoc)
        PartDoc.Close()
        self.oFileSys.DeleteFile(PartDocPath1)
        print self.instance_id
        NewComponent.Name = self.instance_id
        print NewComponent.Name
        print 'EXT COMP ADDED'
        self.catia.ActiveWindow.ActiveViewer.Reframe()

    def change_inst_id(self):

        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        Product2 = to_p.ReferenceProduct
        carm_name = to_p.Name
        carm_name1 = carm_name.replace('_INSTL', '')
        carm_name2 = carm_name1 + '_CARM'
        Product2Products = Product2.Products
        product_forpaste = Product2Products.Item(3)
        product_forpaste.Name = carm_name2
        print product_forpaste.Name

    def access_carm(self):
        """Returns self carm_part"""

        carm_document = self.documents.Item('CA' + self.carm_part_number + '.CATPart')
        carm_part = carm_document.Part
        return carm_part

    def ZZZ_access_annotations(self):

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
        ref_annotation.ModifyVisu()

    def modif_sta_annotation(self, sta_values_fake):

        carm_part = self.access_carm()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        anns = ann_set1.Annotations
        sta_annotation = anns.Item(2)
        ann1text = sta_annotation.Text()
        ann1text_2d = ann1text.Get2dAnnot()
        sta = sta_values_fake[self.copy_from_product - 5]
        # ann1text_value = 'STA ' + sta + '\nLBL 74.3\nWL 294.8\nREF'
        ann1text_value = 'STA ' + sta + '\nREF'
        ann1text_2d.Text = ann1text_value
        print ann1text_value
        sta_annotation.ModifyVisu()

    def select_current_product(self):
        # ICM_1.ApplyWorkMode(2)
        product1 = self.productDocument1.Product
        products1 = product1.Products

        for prod in xrange(1, 5):
            product_to_replace = products1.Item(prod)
            products_to_replace = product_to_replace.Products

            for det in xrange(1, products_to_replace.Count+1):
                product_act_to_replace_nonc = products_to_replace.Item(det)
                if self.cfp_name in str(product_act_to_replace_nonc.Name):
                    return product_act_to_replace_nonc
                else:
                    continue

        for prod in xrange(5, products1.Count+1):
            product_to_replace = products1.Item(prod)
            if self.cfp_name in str(product_to_replace.Name):
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

    def ZZZ_add_geosets(self):
        """Adds Reference Geometry geometrical set"""

        carm_part = self.access_carm()
        geosets = carm_part.HybridBodies
        new_geoset = geosets.add()
        new_geoset.name = 'Reference Geometry1'
        first_gs = geosets.Item(1)
        first_gs.name = 'Renamed'

        # for reference: For IdxSet = 1 To AnnotationSets.Count

    def ZZZ_get_points(self, jd_number):

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
        # print ref_connecter.Name
        # ref_connecter_coordinates_X = ref_connecter.X
        # ref_connecter_coordinates_Y = ref_connecter.Y
        # ref_connecter_coordinates_Z = ref_connecter.Z
        # ref_connecter_new_coordinates.append(ref_connecter_coordinates_X.Value)
        # ref_connecter_new_coordinates.append(ref_connecter_coordinates_Y.Value)
        # ref_connecter_new_coordinates.append(ref_connecter_coordinates_Z.Value)
        # print ref_connecter_new_coordinates
        # ref_connecter_new_coordinates[0] += 300.0
        # print ref_connecter_new_coordinates
        # ref_connecter.SetCoordinates(ref_connecter_new_coordinates)

    def set_parameters(self, sta_value_pairs, size):

        carm_part = self.select_carm_to_paste_data()
        parameters1 = carm_part.Parameters
        ref_param = parameters1.Item('ref_connector_X')
        sta_param = parameters1.Item('sta_connector_X')
        # direct_param = parameters1.Item('view_direction_connector_X')
        print ref_param.Value
        print sta_param.Value
        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        print coord_to_move
        ref_param.Value = coord_to_move + (inch_to_mm(float(size))) - inch_to_mm(0.25)
        sta_param.Value = coord_to_move + inch_to_mm(0.25)
        # direct_param.Value = coord_to_move + (Inch_to_mm(float(size)/2)) + Inch_to_mm(7.0)
        print ref_param.Value
        print sta_param.Value

    def set_standard_parts_params(self, jd_number):

        carm_part = self.select_carm_to_paste_data()
        hole_qty = 0
        parameters1 = carm_part.Parameters
        # for param1 in xrange(1, parameters1.Count):
            # param2 = parameters1.Item(param1)
            # print param2.Name
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

    def copy_bodies_and_paste(self, fastener):
        """Makes copy of fasteners solids and pastes them to the current CARM"""

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.select_current_product()
        selection1.Add(product1)
        selection1.Search(str('(Name = ' + fastener + '*REF-Name = *.*), sel'))
        # selection1.Search(str('Name = ' + fastener + '*REF, sel'))
        try:
            selection1.Copy()
        except:
            pass
        else:
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

    def copy_ref_surface_and_paste(self, size):
        """Makes copy of reference geometry geoset and pastes to current CARM"""

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.select_current_product()
        selection1.Add(product1)
        selection1.Search(str('Name = *' + size + '*IN*REF, sel'))
        selection1.Copy()
        selection2 = self.productDocument1.Selection
        selection2.Clear()
        geoset3 = self.select_carm_to_paste_data()
        hybridBodies1 = geoset3.HybridBodies
        hybridBody1 = hybridBodies1.Item("Reference Geometry")
        selection2.Add(hybridBody1)
        # selection2.visProperties.SetShow(0)
        selection2.PasteSpecial('CATPrtResultWithOutLink')
        geoset3.Update()
        # change visual properties
        selection3 = self.productDocument1.Selection
        selection3.Clear()
        geoset4 = self.select_carm_to_paste_data()
        selection3.Add(geoset4)
        selection3.Search('Name = Reference Geometry, sel')
        selection3.visProperties.SetRealColor(0, 128, 255, 0)
        selection3.visProperties.SetRealOpacity(65, 0)

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

            selection1.Copy()
            self.paste_to_jd(2)

        else:
            if arch:
                selection1.Search(str('(Name = ' + size + '*FAIRING*PLENUM*ARCH*-Name = *CENTERLINE*), sel'))

            else:
                selection1.Search(str('(Name = ' + size + '*FAIRING*PLENUM*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))


            first_elem = selection1.Item2(1)
            first_point = first_elem.Value
            print first_point.Name
            return first_point

    def ZZZ_rename_vectors(self, jd_number):

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
        # annotationFactory1.ActivateTPSView(ann_set1, view_to_activate)

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
            y = inch_to_mm(12)
            z = 0
            if side == 'LH':
                k = 1
            else:
                k = -1
            # addition = k*(Inch_to_mm(float(size)/2))
            addition = k*(inch_to_mm(12.0))
        else:
            y = inch_to_mm(19)
            z = 0
            if side == 'LH':
                k = -1
            else:
                k = 1
            # addition = k*(Inch_to_mm(float(size)/2))
            addition = k*(inch_to_mm(12.0))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 16)

        self.rename_part_body()

        self.hide_last_annotation()
        carm_part.Update()

        # self.activate_top_prod()

    def ZZZ_manage_annotations_visibility(self):

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

    def ZZZ_unhide_std_parts_bodies(self):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        carm_part = self.select_carm_to_paste_data()
        ann_sets = carm_part.AnnotationSets
        std_bodies = carm_part.Bodies
        ann_set1 = ann_sets.Item(1)
        captures1 = ann_set1.Captures
        # for capture in xrange(1, captures1.Count+1):
            # self.access_captures(capture)
        capture1 = captures1.Item(3)
        capture1.DisplayCapture()
        for body in xrange(1, std_bodies.Count+1):
            body1 = std_bodies.Item(body)
            selection1.Add(body1)
        selection1.visProperties.SetShow(0)
        carm_part.Update()

    def ZZZ_unhide_std_parts_bodies1(self):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        carm_part = self.select_carm_to_paste_data()
        std_bodies = carm_part.HybridBodies
        for body in xrange(1, std_bodies.Count+1):
            body1 = std_bodies.Item(body)
            selection1.Add(body1)
        selection1.visProperties.SetShow(0)
        selection1.Clear()
        carm_part.Update()

    def swich_to_part_design(self):

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        part1 = self.select_carm_to_paste_data()
        selection1.Add(part1)
        self.catia.StartWorkbench("PrtCfg")
        print 'part design'
        selection1 = self.productDocument1.Selection
        selection1.Clear()

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
        # info on how to use send keys:
        # wsc = win32com.client.Dispatch("WScript.Shell")
        # wsc.AppActivate("CATIA V5")
        # wsc.SendKeys("c:FrmActivate")
        # wsc.SendKeys("{ENTER}")
        wb = self.workbench_id()
        if wb != 'Assembly':
            self.catia.StartWorkbench("Assembly")
            self.activate_top_prod()

    def workbench_id(self):
        """
        Returns workbench name
        """
        wb_name = self.catia.GetWorkbenchId()
        print str(wb_name)
        return str(wb_name)

    def if_stable(self, sta_values_fake):

        stable_zone = ['0465', '0513', '897', '945', '993', '1041', '1365', '1401', '1401' '1401+0', '1401+48', '1401+96', '1425', '1473', '1521', '1569']
        sta = sta_values_fake[self.copy_from_product - 5]
        if sta in stable_zone:
            return True
        else:
            return False

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
        # carm_part = self.select_carm_to_paste_data()
        documents1 = self.catia.Documents
        partDocument1 = documents1.Item('CA' + self.carm_part_number + '.CATPart')
        cameras = partDocument1.Cameras
        for i in xrange(1, cameras.Count+1):
            camera = cameras.Item(i)
            cam_dict[str(camera.name)] = i
        print cam_dict
        return cam_dict

    def shift_camera(self, sta_value_pairs, size=None):

        # carm_part = self.select_carm_to_paste_data()
        cam_dict = self.map_camera_names()
        documents1 = self.catia.Documents
        partDocument1 = documents1.Item('CA' + self.carm_part_number + '.CATPart')
        cameras = partDocument1.Cameras

        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        coefficient = coord_to_move - (inch_to_mm(717.0 - float(size) / 2))
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
        # sight_direction = [1,0,0]
        # PutUpDirection = [1,1,1]
        # print viewpoint.Viewpoint3D.FocusDistance
        # print viewpoint.Viewpoint3D.Zoom
            vpd = viewpoint.Viewpoint3D
            vpd.PutOrigin(dict_cameras[view])
            vpd.PutSightDirection(dict_sight_directions[view])
            vpd.PutUpDirection(dict_up_directions[view])
        # viewpoint.Viewpoint3D.PutSightDirection(sight_direction)
        # viewpoint.Viewpoint3D.PutUpDirection(sight_direction)

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
        # pointers = capture1.Annotations
        # print pointers.name
        # for pointer in pointers:
            # leader = pointers.Item(pointers.Count)
            # print leader.name
        # anns = ann_set1.Annotations
        # jd1_annotation = anns.Item(4)
        # selection3.Add(jd1_annotation)
        # selection3.visProperties.SetShow(0)


class CarmUpperBin(CarmOmf):

    def __init__(self, carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name, *args):
        super(CarmUpperBin, self).__init__(carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name)
        self.side = side
        self.extention = '\\seed_carm_upr_' + side + '.CATPart'
        for i in args:
            if type(i) is list:
                if len(i) != 0:
                    self.first_elem_in_irm = i[0]
                    self.irm_length = len(i)
                    self.first_elem_in_irm_size = self.first_elem_in_irm[:2]
            else:
                self.plug_value = i

    def select_first_elem_in_irm_product(self):
        # ICM_1.ApplyWorkMode(2)
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

    def change_inst_id_sta(self, sta_values_fake, sta_value_pairs, side, size):

        finish_sta = sta_value((sta_value_pairs[self.copy_from_product - 5] + inch_to_mm(int(size))), self.plug_value)
        start_sta = sta_values_fake[self.copy_from_product - (5 + (self.irm_length - 1))]
        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        instance_id_IRM = 'ECS_UPR-AIR-DIST_INSTL_STA' + start_sta + '-' + finish_sta + '_' + side[0]
        to_p.Name = instance_id_IRM
        print to_p.Name

    def modif_ref_annotation(self, size):

        carm_part = self.access_carm()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        anns = ann_set1.Annotations
        ref_annotation = anns.Item(1)
        ann1text = ref_annotation.Text()
        ann1text_2d = ann1text.Get2dAnnot()
        ann1text_value = size + 'IN OUTBD BIN SUPPORT REF'
        ann1text_2d.Text = ann1text_value
        print ann1text_value
        ref_annotation.ModifyVisu()

    def add_ref_annotation(self, sta_value_pairs, size, side):
        """Adds REF annotation"""

        annot_text = str(size) + 'IN OUTBD BIN SUPPORT REF'
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
        coord_to_move_ref_point = sta_value_pairs[self.copy_from_product - 5] + (inch_to_mm(float(size)) * 0.7)
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
        y = inch_to_mm(30)
        z = 0
        if side == 'LH':
            k = -1
        else:
            k = 1
        addition = k*(inch_to_mm(float(size) / 2))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        text_leaders = ann1text_2d.Leaders
        text_leader1 = text_leaders.Item(1)
        text_leader1.HeadSymbol = 20
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
        coord_to_move_ref_point = sta_value_pairs[self.copy_from_product - 5] + inch_to_mm(0.25)
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
        coord_to_move = sta_value_pairs[self.copy_from_product - 5] - inch_to_mm(2)
        y = inch_to_mm(20)
        z = 0
        if side == 'LH':
            k = -1
        else:
            k = 1
        # addition = k*(Inch_to_mm(float(size)/2))*0
        addition = k*(inch_to_mm(2.15))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        text_leaders = ann1text_2d.Leaders
        text_leader1 = text_leaders.Item(1)
        text_leader1.HeadSymbol = 1
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
        # annotationFactory1 = ann_set1.AnnotationFactory
        ann_set1.ActiveView = view_to_activate
        # annotationFactory1.ActivateTPSView(ann_set1, view_to_activate)

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
            # selection1.Search(str('(Name = ' + size + '*BIN*LIGHT*1_CENTERLINE*), sel'))
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
            if 'ARCH' in self.first_elem_in_irm:
                selection1.Search(str('(Name = ' + size + '*ARCH*PLENUM*UPR*-Name = *CENTERLINE*), sel'))
                print 'ARCH'

            else:
                selection1.Search(str('(Name = ' + size + '*BIN*PLENUM*UPR*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
                print 'NOT ARCH'
            first_elem = selection1.Item2(1)
            first_point = first_elem.Value
            print first_point.Name
            return first_point

    def shift_camera(self, sta_value_pairs, size=None):

        # carm_part = self.select_carm_to_paste_data()
        cam_dict = self.map_camera_names()
        documents1 = self.catia.Documents
        partDocument1 = documents1.Item('CA' + self.carm_part_number + '.CATPart')
        cameras = partDocument1.Cameras

        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        coefficient = coord_to_move - (inch_to_mm(717.0 - float(size) / 2))
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
        # sight_direction = [1,0,0]
        # PutUpDirection = [1,1,1]
        # print viewpoint.Viewpoint3D.FocusDistance
        # print viewpoint.Viewpoint3D.Zoom
            vpd = viewpoint.Viewpoint3D
            vpd.PutOrigin(dict_cameras[view])
            vpd.PutSightDirection(dict_sight_directions[view])
            vpd.PutUpDirection(dict_up_directions[view])
        # viewpoint.Viewpoint3D.PutSightDirection(sight_direction)
        # viewpoint.Viewpoint3D.PutUpDirection(sight_direction)

    def set_parameters(self, sta_value_pairs, size):

        carm_part = self.select_carm_to_paste_data()
        parameters1 = carm_part.Parameters
        ref_param = parameters1.Item('ref_connector_X')
        sta_param = parameters1.Item('sta_connector_X')
        # direct_param = parameters1.Item('view_direction_connector_X')
        print ref_param.Value
        print sta_param.Value
        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        print coord_to_move
        ref_param.Value = coord_to_move + (inch_to_mm(float(size))) - (inch_to_mm(float(size)) * 0.3)
        sta_param.Value = coord_to_move + inch_to_mm(0.25)
        # direct_param.Value = coord_to_move + (Inch_to_mm(float(size)/2)) + Inch_to_mm(7.0)
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
        coord_to_move = sta_value_pairs[self.copy_from_product - 5 - (self.irm_length - 1)] + JD_point_X
        print coord_to_move
        if jd_number == '01':
            y = inch_to_mm(12)
            z = 0
            if side == 'LH':
                k = 1
            else:
                k = -1
            # addition = k*(Inch_to_mm(float(size)/2))
            addition = k * inch_to_mm(12.0)
        else:
            y = inch_to_mm(19)
            z = 0
            if side == 'LH':
                k = -1
            else:
                k = 1
            # addition = k*(Inch_to_mm(float(size)/2))
            addition = k * inch_to_mm(12.0)
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 24)

        self.rename_part_body()

        self.hide_last_annotation()
        carm_part.Update()


class CarmUpperBinNonConstant(CarmOmf):

    def __init__(self, carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name):
        super(CarmUpperBinNonConstant, self).__init__(carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name)
        self.side = side
        self.extention = '\\seed_carm_nonc_' + side + '.CATPart'

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
        coord_to_move_ref_point = sta_value_pairs[self.copy_from_product - 5] + (inch_to_mm(float(size)) / 2.0)
        if side == 'LH':
            coefnt = -1
        else:
            coefnt = 1
        # hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointOnCurveFromDistance(coord_to_move_ref_point, coefnt*1878.360480, 7493.000000)
        # geoset2.AppendHybridShape(hybridShapePointCoord1)
        # carm_part.Update()
        points = geoset2.HybridShapes
        wb = str(self.workbench_id())
        if wb != 'PrtCfg':
            self.swich_to_part_design()
        reference2 = carm_part.CreateReferenceFromObject(points.Item('point_direction'))
        userSurface2 = userSurfaces1.Generate(reference2)
        hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointOnCurveFromDistance(userSurface2, coord_to_move_ref_point, True)
        geoset2.AppendHybridShape(hybridShapePointCoord1)
        reference1 = carm_part.CreateReferenceFromObject(points.Item(points.Count))
        carm_part.Update()
        r = points.Item(points.Count)
        print r.Name
        userSurface1 = userSurfaces1.Generate(reference1)
        annotationFactory1 = ann_set1.AnnotationFactory
        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        y = inch_to_mm(28)
        z = 0
        if side == 'LH':
            k = -1
        else:
            k = 1
        addition = k*(inch_to_mm(float(size) / 2))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        text_leaders = ann1text_2d.Leaders
        text_leader1 = text_leaders.Item(1)
        text_leader1.HeadSymbol = 20
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
        # hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(coord_to_move_ref_point, coefnt*1878.360480, 7493.000000)
        # geoset2.AppendHybridShape(hybridShapePointCoord1)
        # carm_part.Update()
        points = geoset2.HybridShapes
        wb = str(self.workbench_id())
        if wb != 'PrtCfg':
            self.swich_to_part_design()
        reference2 = carm_part.CreateReferenceFromObject(points.Item('point_direction'))
        userSurface2 = userSurfaces1.Generate(reference2)
        hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointOnCurveFromDistance(userSurface2, coord_to_move_ref_point, True)
        geoset2.AppendHybridShape(hybridShapePointCoord1)
        reference1 = carm_part.CreateReferenceFromObject(points.Item(points.Count))
        carm_part.Update()
        userSurface1 = userSurfaces1.Generate(reference1)
        annotationFactory1 = ann_set1.AnnotationFactory
        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        y = inch_to_mm(22)
        z = 0
        if side == 'LH':
            k = -1
        else:
            k = 1
        addition = k*(inch_to_mm(2.15))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        text_leaders = ann1text_2d.Leaders
        text_leader1 = text_leaders.Item(1)
        text_leader1.HeadSymbol = 1
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
        # annotationFactory1 = ann_set1.AnnotationFactory
        ann_set1.ActiveView = view_to_activate
        # annotationFactory1.ActivateTPSView(ann_set1, view_to_activate)

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
            JD_point = self.copy_jd1_fcm10f5cps05wh_and_paste(size, 'find_point')
        else:
            JD_point = self.copy_jd2_bacs12fa3k3_and_paste(size, arch, 'find_point')

        JD_point_coord_X = JD_point.X
        JD_point_X = JD_point_coord_X.Value
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
        coord_to_move = inch_to_mm(457.917) - (-1 * sta_value_pairs[self.copy_from_product - 5])
        # coord_to_move = Inch_to_mm(457.917) - (-1 * sta_value_pairs[self.copy_from_product - 5]) - JD_point_X - Inch_to_mm(float(size))
        if jd_number == '01':
            y = inch_to_mm(12)
            z = 0
            if side == 'LH':
                k = 1
            else:
                k = -1
            # addition = k*(Inch_to_mm(24))
        else:
            y = inch_to_mm(-12)
            z = 0
            if side == 'LH':
                k = -1
            else:
                k = 1
            # addition = k*(Inch_to_mm(24.0))

        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 16)

        self.rename_part_body()

        self.hide_last_annotation()
        carm_part.Update()

        # self.activate_top_prod()

    def set_parameters(self, sta_value_pairs, size):

        carm_part = self.select_carm_to_paste_data()
        parameters1 = carm_part.Parameters
        ref_param = parameters1.Item('ref_connector_X')
        sta_param = parameters1.Item('sta_connector_X')
        # direct_param = parameters1.Item('view_direction_connector_X')
        print ref_param.Value
        print sta_param.Value
        coord_to_move = -1 * (sta_value_pairs[self.copy_from_product - 5])
        print coord_to_move
        sta_param.Value = coord_to_move + (inch_to_mm(float(size))) - inch_to_mm(0.5)
        ref_param.Value = coord_to_move
        print ref_param.Value
        print sta_param.Value

    def shift_camera(self, sta_value_pairs, size=None):

        cam_dict = self.map_camera_names()
        documents1 = self.catia.Documents
        partDocument1 = documents1.Item('CA' + self.carm_part_number + '.CATPart')
        cameras = partDocument1.Cameras
        coord_to_move = inch_to_mm(457.917) - (-1 * sta_value_pairs[self.copy_from_product - 5])
        coefficient = coord_to_move - inch_to_mm(457.917)

        print coefficient
        if self.side == 'LH':
            all_annotations_origin = [13315.510742+coefficient, 696.444824, 9577.463867]
            engineering_definition_release_origin = [13315.510742+coefficient, 696.444824, 9577.463867]
            reference_geometry_origin = [10226.757813+coefficient, 1628.633789, 10644.860352]
            upper_plenum_fasteners_origin = [10692.801758+coefficient, -6077.0, 8673.823242]
            upper_plenum_spud_fasteners_origin = [11296.463867+coefficient, 1104.921265, 10296.697266]
            all_annotations_sight_direction = [-0.57735, -0.57735, -0.57735]
            all_annotations_up_direction = [-0.4082489, -0.408248, 0.816497]
            engineering_definition_sight_direction = [-0.57735, -0.57735, -0.57735]
            engineering_definition_up_direction = [-0.408248, -0.408248, 0.816497]
            reference_geometry_sight_direction = [-0.05904, -0.674833, -0.735605]
            reference_geometry_up_direction = [-0.064112, -0.735605, 0.67741]
            upper_plenum_fasteners_sight_direction = [0.080368, 0.949164, -0.30435]
            upper_plenum_fasteners_up_direction = [0.026448, 0.3032, 0.95256]
            upper_plenum_spud_fasteners_sight_direction = [-0.05904, -0.674833, -0.735605]
            upper_plenum_spud_fasteners_up_direction = [-0.064112, -0.735605, 0.67741]

        else:
            all_annotations_origin = [8824.277344+coefficient, -1230.204712, 9511.789063]
            engineering_definition_release_origin = [9241.155273+coefficient, -869.085632, 9268.230469]
            reference_geometry_origin = [11314.335938+coefficient, -649.952087, 9972.473633]
            upper_plenum_fasteners_origin = [10815.461914+coefficient, 4949.903809, 8104.419922]
            upper_plenum_spud_fasteners_origin = [11351.40918+coefficient, -945.137512, 10264.790039]
            all_annotations_sight_direction = [0.507266, 0.688048, -0.518913]
            all_annotations_up_direction = [0.304368, 0.420292, 0.854818]
            engineering_definition_sight_direction = [0.507265, 0.688048, -0.518913]
            engineering_definition_up_direction = [0.304368, 0.420292, 0.854818]
            reference_geometry_sight_direction = [-0.05904, 0.674833, -0.735605]
            reference_geometry_up_direction = [-0.064112, 0.732806, 0.677411]
            upper_plenum_fasteners_sight_direction = [0.085044, -0.960193, -0.266077]
            upper_plenum_fasteners_up_direction = [0.021026, -0.265255, 0.963949]
            upper_plenum_spud_fasteners_sight_direction = [-0.05904, 0.674833, -0.735605]
            upper_plenum_spud_fasteners_up_direction = [-0.064112, 0.732806, 0.677411]

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
        # sight_direction = [1,0,0]
        # PutUpDirection = [1,1,1]
        # print viewpoint.Viewpoint3D.FocusDistance
        # print viewpoint.Viewpoint3D.Zoom
            vpd = viewpoint.Viewpoint3D
            vpd.PutOrigin(dict_cameras[view])
            vpd.PutSightDirection(dict_sight_directions[view])
            vpd.PutUpDirection(dict_up_directions[view])


        # viewpoint.Viewpoint3D.PutSightDirection(sight_direction)
        # viewpoint.Viewpoint3D.PutUpDirection(sight_direction)


class CarmLowerBin(CarmUpperBin):

    def __init__(self, carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name, state, *args):
        super(CarmLowerBin, self).__init__(carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name, *args)
        self.side = side
        self.extention = '\\seed_carm_lwr_' + side + '.CATPart'

        count = 0
        for i in args:
            if type(i) is list:
                if len(i) != 0:
                    self.first_elem_in_irm_new = i[0]
            else:
                self.plug_value = i
        #if state == 'final':
        for i in args:
            if type(i) is not list:
                break
            else:
                for n in xrange(7):
                    count += 1
                    if '24' in i[n]:
                        if len(i) == count:
                            print '24 found'
                            self.first_elem_in_irm = i[n]
                            count = 0
                            break
                        else:
                            continue
                    else:
                        self.first_elem_in_irm = i[n]
                        print 'first_elem_is_' + str(self.first_elem_in_irm)
                        break
            self.irm_components = i
            self.irm_length = len(i)
            self.first_elem_in_irm_size = self.first_elem_in_irm[:2]

    def copy_jd1_BACS12FA3K20_and_paste(self, size, type_of_geometry='points'):

        selection1 = self.productDocument1.Selection
        selection1.Clear()


        if type_of_geometry == 'points':
            product1 = self.select_current_product()
            selection1.Add(product1)
            selection1.Search(str('(Name = ' + size + '*BIN*PLENUM*LWR*BACI12AG3UCM2*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
            selection1.Copy()
            self.paste_to_jd(1)
        else:
            product1 = self.select_first_elem_in_irm_product()
            selection1.Add(product1)
            selection1.Search(str('(Name = ' + size + '*BIN*PLENUM*LWR*BACI12AG3UCM2*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
            first_elem = selection1.Item2(1)
            first_point = first_elem.Value
            print first_point.Name
            return first_point

    def copy_jd2_bacs12fa3k3_and_paste_1(self, size, type_of_geometry='points'):

        selection1 = self.productDocument1.Selection
        selection1.Clear()

        if type_of_geometry == 'points':
            product1 = self.select_current_product()
            selection1.Add(product1)
            selection1.Search(str('(Name = ' + size + '*BIN*PLENUM*LWR*BACI12AK3CM07*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
            selection1.Copy()
            self.paste_to_jd(2)

        elif type_of_geometry == 'last_point':
            product1 = self.select_first_elem_in_irm_product()
            selection1.Add(product1)
            selection1.Search(str('(Name = ' + size + '*BIN*PLENUM*LWR*BACI12AK3CM07*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
            first_elem = selection1.Item2(selection1.Count2)
            first_point = first_elem.Value
            print first_point.Name
            return first_point

        else:
            product1 = self.select_first_elem_in_irm_product()
            selection1.Add(product1)
            selection1.Search(str('(Name = ' + size + '*BIN*PLENUM*LWR*BACI12AK3CM07*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
            first_elem = selection1.Item2(1)
            first_point = first_elem.Value
            print first_point.Name
            return first_point

    def copy_jd3_BACS12FA3K12_and_paste(self, size, type_of_geometry='points'):

        selection1 = self.productDocument1.Selection
        selection1.Clear()

        if type_of_geometry == 'points':
            product1 = self.select_current_product()
            selection1.Add(product1)
            selection1.Search(str('(Name = ' + size + '*NOZZLE*LOWER*BACI12AH5U375*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
            selection1.Copy()
            self.paste_to_jd(3)
        else:
            product1 = self.select_first_elem_in_irm_product()
            selection1.Add(product1)
            selection1.Search(str('(Name = ' + size + '*NOZZLE*LOWER*BACI12AH5U375*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
            first_elem = selection1.Item2(1)
            first_point = first_elem.Value
            print first_point.Name
            return first_point

    def copy_jd4_bacs12fa3k3_and_paste_2(self, size, type_of_geometry='points'):

        selection1 = self.productDocument1.Selection
        selection1.Clear()

        if type_of_geometry == 'points':
            product1 = self.select_current_product()
            selection1.Add(product1)
            selection1.Search(str('(Name = ' + size + '*OB*BIN*END*FRAME*ECS*NOZZLE*BACI12AK3CM07*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
            selection1.Copy()
            self.paste_to_jd(4)
        else:
            product1 = self.select_first_elem_in_irm_product()
            selection1.Add(product1)
            selection1.Search(str('(Name = ' + size + '*OB*BIN*END*FRAME*ECS*NOZZLE*BACI12AK3CM07*-(Name = *CENTERLINE*+Name = *ARCH*+Name = *SEC*47*)), sel'))
            first_elem = selection1.Item2(1)
            first_point = first_elem.Value
            print first_point.Name
            return first_point

    def set_parameters(self, sta_value_pairs, size):

        carm_part = self.select_carm_to_paste_data()
        #carm_part = self.access_carm()
        parameters1 = carm_part.Parameters
        FL2_X_param = parameters1.Item('FL2_X')
        FL3_X_param = parameters1.Item('FL3_X')
        FL4_X_param = parameters1.Item('FL4_X')
        FL5_X_param = parameters1.Item('FL5_X')
        fir_tree_param = parameters1.Item('156-00066a')
        bacs_param = parameters1.Item('BACS38K2a')
        FL2_X_param_offset = inch_to_mm(4.0606)
        FL3_X_param_offset = inch_to_mm(2.1582)
        FL4_X_param_offset = inch_to_mm(3.0059)
        FL5_X_param_offset = inch_to_mm(2.4261)
        fir_tree_param_offset = inch_to_mm(0.4353)
        bacs_param_offset = inch_to_mm(1.4077)
        anchor_point = self.copy_jd2_bacs12fa3k3_and_paste_1(size, 'find_point')
        anchor_point_coord_X = anchor_point.X
        anchor_point_X = anchor_point_coord_X.Value
        coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        print 'SET_PARAMETERS ' + str(coord_to_move)
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
        coefficient = coord_to_move - (inch_to_mm(717.0 - float(size) / 2))
        print coefficient

        if self.side == 'LH':
            engineering_definition_release_origin = [12313.794922 + coefficient + 4000.0, -6951.88916, 9594.439453]
            all_annotations_origin = [10896.336914 + coefficient + 4000.0, -7317.958984, 10955.158203]
            reference_geometry_origin = [15983.475586 + coefficient + 4000.0, 3519.238281, 7186.692383]
            lower_plenum_downer_strap_origin = [16046.920898 + coefficient + 4000.0, -7975.490723, 6932.231445]
            upper_downer_strap_origin = [15985.112305 + coefficient + 4000.0, -7302.797852, 11938.787109]
            lower_plenum_fastener_jd01_origin = [15939.419922 + coefficient + 4000.0, -9042.191406, 7099.609375]
            lower_plenum_fastener_jd02_origin = [15958.158203 + coefficient + 4000.0, -6916.774902, 11606.900391]
            sidewall_nozzle_fastener_jd03_origin = [16082.665039 + coefficient + 4000.0, 4481.557617, 4473.383301]
            sidewall_nozzle_fastener_jd04_origin = [15905.139648 + coefficient + 4000.0, -8642.466797, 6985.853516]

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
            lower_plenum_fastener_jd01_sight_direction = [0, 1, 0]
            lower_plenum_fastener_jd01_up_direction = [0, 0, 1]
            lower_plenum_fastener_jd02_sight_direction = [0, 0.726964, -0.686676]
            lower_plenum_fastener_jd02_up_direction = [0, 0.686676, 0.726964]
            sidewall_nozzle_fastener_jd03_sight_direction = [-0.011397, -0.937411, 0.348038]
            sidewall_nozzle_fastener_jd03_up_direction = [0, 0.348079, 0.937464]
            sidewall_nozzle_fastener_jd04_sight_direction = [0, 1, 0]
            sidewall_nozzle_fastener_jd04_up_direction = [0, 0, 1]

        else:
            engineering_definition_release_origin = [22140.060547 + coefficient + 4000.0, 8147.486816, 13061.50293]
            all_annotations_origin = [22481.875 + coefficient + 4000.0, 8442.293945, 13728.501953]
            reference_geometry_origin = [15888.384766 + coefficient + 4000.0, -10043.158203, 7162.20752]
            lower_plenum_downer_strap_origin = [16132.829102 + coefficient + 4000.0, 4108.947266, 7107.040039]
            upper_downer_strap_origin = [15898.631836 + coefficient + 4000.0, -2836.321289, 11428.59668]
            lower_plenum_fastener_jd01_origin = [15984.461914 + coefficient + 4000.0, 9607.047852, 6911.10498]
            lower_plenum_fastener_jd02_origin = [16045.074219 + coefficient + 4000.0, 7321.30957, 11803.600586]
            sidewall_nozzle_fastener_jd03_origin = [16035.198242 + coefficient + 4000.0, -4300.676758, 4722.25293]
            sidewall_nozzle_fastener_jd04_origin = [16086.536133 + coefficient + 4000.0, 4363.338867, 7020.689941]

            engineering_definition_sight_direction = [-0.57735, -0.57735, -0.57735]
            engineering_definition_up_direction = [-0.408248, -0.408248, 0.816497]
            all_annotations_sight_direction = [-0.57735, -0.57735, -0.57735]
            all_annotations_up_direction = [-0.411367, -0.405122, 0.816489]
            reference_geometry_sight_direction = [0, 1, 0]
            reference_geometry_up_direction = [0, 0, 1]
            lower_plenum_downer_strap_sight_direction = [0, -1, 0]
            lower_plenum_downer_strap_up_direction = [0, 0, 1]
            upper_downer_strap_sight_direction = [0, 0.734878, -0.6782]
            upper_downer_strap_up_direction = [0, 0.6782, 0.734878]
            lower_plenum_fastener_jd01_sight_direction = [0, -1, 0]
            lower_plenum_fastener_jd01_up_direction = [0, 0, 1]
            lower_plenum_fastener_jd02_sight_direction = [0, -0.726964, -0.686676]
            lower_plenum_fastener_jd02_up_direction = [0, -0.686676, 0.726964]
            sidewall_nozzle_fastener_jd03_sight_direction = [0, 0.945367, 0.325971]
            sidewall_nozzle_fastener_jd03_up_direction = [0, -0.325973, 0.945379]
            sidewall_nozzle_fastener_jd04_sight_direction = [0, -1, 0]
            sidewall_nozzle_fastener_jd04_up_direction = [0, 0, 1]

        view_origins = [all_annotations_origin, engineering_definition_release_origin, reference_geometry_origin, lower_plenum_downer_strap_origin, upper_downer_strap_origin, lower_plenum_fastener_jd01_origin, lower_plenum_fastener_jd02_origin, sidewall_nozzle_fastener_jd03_origin, sidewall_nozzle_fastener_jd04_origin]
        viewpoints = ['ALL ANNOTATION', 'ENGINEERING DEFINITION RELEASE', 'REFERENCE GEOMETRY', 'LOWER PLENUM DOWNER STRAP', 'UPPER DOWNER STRAP', 'LOWER PLENUM FASTENER JD01', 'LOWER PLENUM FASTENER JD02', 'SIDEWALL NOZZLE FASTENER JD03', 'SIDEWALL NOZZLE FASTENER JD04']
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
        print to_p.Name
        if '24' in self.first_elem_in_irm_new:
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

    def change_inst_id(self):

        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        Product2 = to_p.ReferenceProduct
        Product2Products = Product2.Products
        carm_name = to_p.Name
        carm_name1 = carm_name.replace('_INSTL', '')
        carm_name2 = carm_name1 + '_CARM'
        if '24' in self.first_elem_in_irm_new:
            product_forpaste = Product2Products.Item(3)
        else:
            product_forpaste = Product2Products.Item(4)
        product_forpaste.Name = carm_name2
        print product_forpaste.Name

    def change_inst_id_sta(self, sta_values_fake, sta_value_pairs, side, size):

        finish_sta = sta_value((sta_value_pairs[self.copy_from_product - 5] + inch_to_mm(int(size))), self.plug_value)
        start_sta = sta_values_fake[self.copy_from_product - (5 + (self.irm_length - 1))]
        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        # Product2Products = Product2.Products
        # product_forpaste = Product2Products.Item(3)
        instance_id_IRM = 'ECS_LWR-AIR-DIST_INSTL_STA' + start_sta + '-' + finish_sta + '_' + side[0]
        to_p.Name = instance_id_IRM
        print to_p.Name

    def copy_bodies_and_paste(self, fastener):
        """Makes copy of fasteners solids and pastes them to the current CARM"""

        selection1 = self.productDocument1.Selection
        selection1.Clear()
        product1 = self.select_current_product()
        selection1.Add(product1)
        # selection1.Search(str('(Name = ' + fastener + '*REF-Name = *.*), sel'))
        selection1.Search(str('Name = ' + fastener + '*REF, sel'))
        try:
            selection1.Copy()
        except:
            pass
        else:
            selection2 = self.productDocument1.Selection
            selection2.Clear()
            part2 = self.select_carm_to_paste_data()
            selection2.Add(part2)
            selection2.PasteSpecial('CATPrtResultWithOutLink')
            part2.Update()

    def add_ref_annotation(self, sta_value_pairs, size, side):

        """Adds REF annotation"""

        annot_text = str(size) + 'IN OUTBD BIN SUPPORT REF'
        carm_part = self.access_carm()
        self.activate_view(3)
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        userSurfaces1 = carm_part.UserSurfaces
        geosets = carm_part.HybridBodies
        geoset1 = geosets.Item('Construction Geometry (REF)')
        geosets1 = geoset1.HybridBodies
        geoset2 = geosets1.Item('Misc Construction Geometry')
        hybridShapeFactory1 = carm_part.HybridShapeFactory
        coord_to_move_ref_point = sta_value_pairs[self.copy_from_product - 5] + (inch_to_mm(float(size)) * 0.7)
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
        y = inch_to_mm(60)
        z = 0
        if side == 'LH':
            k = -1
        else:
            k = 1
        addition = k*(inch_to_mm(float(size) / 2))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        text_leaders = ann1text_2d.Leaders
        text_leader1 = text_leaders.Item(1)
        text_leader1.HeadSymbol = 20
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
        self.activate_view(3)
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        userSurfaces1 = carm_part.UserSurfaces
        geosets = carm_part.HybridBodies
        geoset1 = geosets.Item('Construction Geometry (REF)')
        geosets1 = geoset1.HybridBodies
        geoset2 = geosets1.Item('Misc Construction Geometry')
        points = geoset2.HybridShapes
        hybridShapeFactory1 = carm_part.HybridShapeFactory
        coord_to_move_ref_point = sta_value_pairs[self.copy_from_product - 5] + inch_to_mm(0.25)
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
        coord_to_move = sta_value_pairs[self.copy_from_product - 5] - inch_to_mm(2)
        y = inch_to_mm(40)
        z = 0
        if side == 'LH':
            k = -1
        else:
            k = 1
        # addition = k*(Inch_to_mm(float(size)/2))*0
        addition = k*(inch_to_mm(2.15))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        text_leaders = ann1text_2d.Leaders
        text_leader1 = text_leaders.Item(1)
        text_leader1.HeadSymbol = 1
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 24)
        print ann1text_2d.AnchorPosition
        ann1text_2d.AnchorPosition = 6
        print ann1text_2d.AnchorPosition
        ann1text_2d.FrameType = 3
        self.rename_part_body()
        self.hide_last_annotation()
        carm_part.Update()

    def add_jd_annotation(self, jd_number, sta_value_pairs, size, side, arch):
        """Adds JOINT DEFINITION XX annotation"""

        annot_text = 'JOINT DEFINITION ' + jd_number
        carm_part = self.access_carm()
        # self.activate_view(jd_number)
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        userSurfaces1 = carm_part.UserSurfaces
        geosets = carm_part.HybridBodies
        geoset1 = geosets.Item('Joint Definitions')
        geosets1 = geoset1.HybridBodies
        geoset2 = geosets1.Item('Joint Definition ' + jd_number)
        points = geoset2.HybridShapes
        if '1' in str(jd_number):
            JD_point = self.copy_jd1_BACS12FA3K20_and_paste(self.first_elem_in_irm_size, 'find_point')
        elif '2' in str(jd_number):
            JD_point = self.copy_jd2_bacs12fa3k3_and_paste_1(self.first_elem_in_irm_size, 'find_point')
        elif '3' in str(jd_number):
            JD_point = self.copy_jd3_BACS12FA3K12_and_paste(self.first_elem_in_irm_size, 'find_point')
        else:
            JD_point = self.copy_jd4_bacs12fa3k3_and_paste_2(self.first_elem_in_irm_size, 'find_point')
        JD_point_coord_X = JD_point.X
        JD_point_X = JD_point_coord_X.Value
        print JD_point_X
        wb = str(self.workbench_id())
        if wb != 'PrtCfg':
            self.swich_to_part_design()
        selection1 = self.productDocument1.Selection
        selection1.Clear()
        reference1 = carm_part.CreateReferenceFromObject(points.Item(1))
        userSurface1 = userSurfaces1.Generate(reference1)

        for point in xrange(2, points.Count+1):
            reference2 = carm_part.CreateReferenceFromObject(points.Item(point))
            print reference2.name
            userSurface1.AddReference(reference2)

        annotationFactory1 = ann_set1.AnnotationFactory
        coord_to_move = sta_value_pairs[self.copy_from_product - 5 - (self.irm_length - 1)] + JD_point_X
        print coord_to_move

        y = inch_to_mm(12)
        z = 0
        if side == 'LH':
            k = 1
        else:
            k = -1
        if '3' in str(jd_number):
            k = k * (-1)
        addition = k * inch_to_mm(12.0)

        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 24)
        self.rename_part_body()
        self.hide_last_annotation()
        carm_part.Update()

    def create_jd_vectors(self, jd_number):

        points_in_geoset = self.points_ammount(jd_number)
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
        Orientation = True
        if jd_number == 1:
            Orientation = False
        elif jd_number == 4 and self.side == 'RH':
            Orientation = False
        if jd_number == 4:
            for point in range(1, points_in_geoset + 1):
                hybridShapePointCenter1 = hybridShapes1.Item(point)
                reference1 = part1.CreateReferenceFromObject(hybridShapePointCenter1)
                hybridShapeLinePtDir1 = hybridShapeFactory1.AddNewLinePtDir(reference1, hybridShapeDirection1, 0.000000, 25.400000, Orientation)
                Orientation = not Orientation
                hybridBody2.AppendHybridShape(hybridShapeLinePtDir1)
                hybridShapeLinePtDir1.Name = 'FIDV_0' + str(jd_number)
                #part1.Update()
        else:
            hybridShapeLinePtDir1 = hybridShapeFactory1.AddNewLinePtDir(reference1, hybridShapeDirection1, 0.000000, 25.400000, Orientation)
            hybridBody2.AppendHybridShape(hybridShapeLinePtDir1)
            hybridShapeLinePtDir1.Name = 'FIDV_0' + str(jd_number)
            #part1.Update()

    def points_ammount(self, jd_number):

        elements_in_geoset = 0
        part1 = self.select_carm_to_paste_data()
        hybridBodies1 = part1.HybridBodies
        hybridBody1 = hybridBodies1.Item("Joint Definitions")
        hybridBodies2 = hybridBody1.HybridBodies
        hybridBody2 = hybridBodies2.Item("Joint Definition" + ' 0' + str(jd_number))
        hybridShapes1 = hybridBody2.HybridShapes
        for i in range(1, hybridShapes1.Count + 1):
            elements_in_geoset += 1
        return elements_in_geoset

    def modif_lwr_strap_annotation(self):

        carm_part = self.access_carm()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        anns = ann_set1.Annotations
        sta_annotation = anns.Item(7)
        ann1text = sta_annotation.Text()
        ann1text_2d = ann1text.Get2dAnnot()
        number_of_straps = 0
        print self.irm_components
        for component in self.irm_components:
            if '24' in str(component):
                number_of_straps += 1
            else:
                number_of_straps += 2
        #ann1text_value = 'STA ' + sta + '\nLBL 74.3\nWL 294.8\nREF'
        ann1text_value = 'BACS38K2 TYPICAL' + '\n' + str(number_of_straps) + ' PLACES'
        ann1text_2d.Text = ann1text_value
        print ann1text_value
        sta_annotation.ModifyVisu()

    def modif_upr_strap_annotation(self):

        carm_part = self.access_carm()
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        anns = ann_set1.Annotations
        sta_annotation = anns.Item(1)
        ann1text = sta_annotation.Text()
        ann1text_2d = ann1text.Get2dAnnot()
        number_of_straps = 0
        for component in self.irm_components:
            if '24' in str(component):
                number_of_straps += 1
            else:
                number_of_straps += 2
        ann1text_value = '156-00066 TYPICAL' + '\n' + str(number_of_straps * 2) + ' PLACES'
        ann1text_2d.Text = ann1text_value
        print ann1text_value
        sta_annotation.ModifyVisu()


class CarmOmfNonConstant(CarmUpperBin):

    def __init__(self, carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name, *args):
        super(CarmUpperBin, self).__init__(carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name)
        self.side = side
        self.extention = '\\seed_carm_nonc_irm_' + side + '.CATPart'
        for i in args:
            if len(i) != 0:
                self.first_elem_in_irm = i[0]
                self.irm_length = len(i)
                self.first_elem_in_irm_size = self.first_elem_in_irm[:2]

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

        if jd_number == '01':
            y = inch_to_mm(12)
            z = 0
            if side == 'LH':
                k = 1
            else:
                k = -1
            #addition = k*(Inch_to_mm(float(size)/2))
            addition = k * inch_to_mm(12.0)
        else:
            y = inch_to_mm(-10)
            z = 0
            if side == 'LH':
                k = -1
            else:
                k = 1
            #addition = k*(Inch_to_mm(float(size)/2))
            addition = k * inch_to_mm(12.0)

        coord_to_move = inch_to_mm(457.917) + sta_value_pairs[self.copy_from_product - 5] - (inch_to_mm(float(size))) + k * JD_point_X
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 24)

        self.rename_part_body()

        self.hide_last_annotation()
        carm_part.Update()

        #self.activate_top_prod()

    def add_ref_annotation(self, sta_value_pairs, size, side):
        """Adds REF annotation"""

        annot_text = str(size) + 'IN OUTBD BIN SUPPORT REF'
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
        coord_to_move_ref_point = -1*(sta_value_pairs[self.copy_from_product - 5]) + (inch_to_mm(float(size)) * 0.3)
        print 'ref_point: ' + str(coord_to_move_ref_point)
        points = geoset2.HybridShapes
        wb = str(self.workbench_id())
        if wb != 'PrtCfg':
            self.swich_to_part_design()
        reference2 = carm_part.CreateReferenceFromObject(points.Item('point_direction'))
        print reference2.Name
        userSurface2 = userSurfaces1.Generate(reference2)
        hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointOnCurveFromDistance(reference2, coord_to_move_ref_point, False)
        geoset2.AppendHybridShape(hybridShapePointCoord1)
        reference1 = carm_part.CreateReferenceFromObject(points.Item(points.Count))
        carm_part.Update()
        r = points.Item(points.Count)
        print r.Name
        userSurface1 = userSurfaces1.Generate(reference1)
        annotationFactory1 = ann_set1.AnnotationFactory
        coord_to_move = inch_to_mm(457.917) - coord_to_move_ref_point
        y = inch_to_mm(-3.0)
        z = 0
        if side == 'LH':
            k = -1
        else:
            k = 1
        addition = -1*k*(inch_to_mm(float(size) / 2))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        text_leaders = ann1text_2d.Leaders
        text_leader1 = text_leaders.Item(1)
        text_leader1.HeadSymbol = 20
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 20)
        self.rename_part_body()
        self.hide_last_annotation()
        carm_part.Update()

    def add_sta_annotation(self, sta_value_pairs, sta_values_fake, size, side):
        """Adds STA annotation"""

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
        coord_to_move_ref_point = -1*(sta_value_pairs[self.copy_from_product - 5]) + (inch_to_mm(float(size))) - inch_to_mm(0.5)
        points = geoset2.HybridShapes
        wb = str(self.workbench_id())
        if wb != 'PrtCfg':
            self.swich_to_part_design()
        reference2 = carm_part.CreateReferenceFromObject(points.Item('point_direction'))
        userSurface2 = userSurfaces1.Generate(reference2)
        hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointOnCurveFromDistance(reference2, coord_to_move_ref_point, False)
        geoset2.AppendHybridShape(hybridShapePointCoord1)
        reference1 = carm_part.CreateReferenceFromObject(points.Item(points.Count))
        carm_part.Update()
        userSurface1 = userSurfaces1.Generate(reference1)
        annotationFactory1 = ann_set1.AnnotationFactory
        coord_to_move = inch_to_mm(457.917) - coord_to_move_ref_point
        y = inch_to_mm(-10.0)
        z = 0
        if side == 'LH':
            k = -1
        else:
            k = 1
        addition = k*(inch_to_mm(2.15))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        text_leaders = ann1text_2d.Leaders
        text_leader1 = text_leaders.Item(1)
        text_leader1.HeadSymbol = 1
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 20)
        ann1text_2d.AnchorPosition = 6
        ann1text_2d.FrameType = 3
        self.rename_part_body()
        self.hide_last_annotation()
        carm_part.Update()

    def set_parameters(self, sta_value_pairs, size):

        carm_part = self.select_carm_to_paste_data()
        parameters1 = carm_part.Parameters
        ref_param = parameters1.Item('ref_connector_X')
        sta_param = parameters1.Item('sta_connector_X')
        #direct_param = parameters1.Item('view_direction_connector_X')
        print ref_param.Value
        print sta_param.Value
        coord_to_move = -1 * (sta_value_pairs[self.copy_from_product - 5])
        print coord_to_move
        sta_param.Value = coord_to_move + (inch_to_mm(float(size))) - inch_to_mm(0.5)
        ref_param.Value = coord_to_move
        #direct_param.Value = coord_to_move + (Inch_to_mm(float(size)/2)) + Inch_to_mm(7.0)
        print ref_param.Value
        print sta_param.Value

    def change_inst_id_sta(self, sta_values_fake, side):

        start_sta = sta_values_fake[self.copy_from_product - 5]
        #actual sta:
        #finish_sta = int((sta_values_fake[self.copy_from_product - (5 + (self.irm_length - 1))])[1:]) + int(self.first_elem_in_irm_size)
        finish_sta = 465
        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        Product2 = to_p.ReferenceProduct
        #Product2Products = Product2.Products
        #product_forpaste = Product2Products.Item(3)
        instance_id_IRM = 'ECS_UPR-AIR-DIST_INSTL_STA' + start_sta + '-0' + str(finish_sta) + '_' + side[0]
        to_p.Name = instance_id_IRM
        print to_p.Name

    def shift_camera(self, sta_value_pairs, size=None):

        #carm_part = self.select_carm_to_paste_data()
        cam_dict = self.map_camera_names()
        documents1 = self.catia.Documents
        partDocument1 = documents1.Item('CA' + self.carm_part_number + '.CATPart')
        cameras = partDocument1.Cameras


        coord_to_move = inch_to_mm(457.917) - (-1 * sta_value_pairs[self.copy_from_product - 5])

        coefficient = coord_to_move - inch_to_mm(457.917)

        print coefficient
        if self.side == 'LH':
            all_annotations_origin = [13854.880859+coefficient, 4417.077637, 12727.319336]
            engineering_definition_release_origin = [13854.880859+coefficient, 4417.077637, 12727.319336]
            reference_geometry_origin = [9994.436523+coefficient, 1648.96167, 10644.861328]
            upper_plenum_fasteners_origin = [9306.952148+coefficient, -7007.398926, 8625.181641]
            upper_plenum_spud_fasteners_origin = [10030.945313+coefficient, 3074.359619, 11994.473633]
            all_annotations_sight_direction = [-0.57735, -0.57735, -0.57735]
            all_annotations_up_direction = [-0.4082489, -0.408248, 0.816497]
            engineering_definition_sight_direction = [-0.57735, -0.57735, -0.57735]
            engineering_definition_up_direction = [-0.408248, -0.408248, 0.816497]
            reference_geometry_sight_direction = [-0.05904, -0.674833, -0.735605]
            reference_geometry_up_direction = [-0.064112, -0.732806, 0.677411]
            upper_plenum_fasteners_sight_direction = [0.080368, 0.949164, -0.30435]
            upper_plenum_fasteners_up_direction = [0.026448, 0.3032, 0.95256]
            upper_plenum_spud_fasteners_sight_direction = [-0.05904, -0.674833, -0.735605]
            upper_plenum_spud_fasteners_up_direction = [-0.064112, -0.732806, 0.677411]

        else:
            all_annotations_origin = [5029.863281+coefficient, -2199.449463, 11174.419922]
            engineering_definition_release_origin = [5029.863281+coefficient, -2199.449463, 11174.419922]
            reference_geometry_origin = [10082.896484+coefficient, -2005.115234, 11119.094727]
            upper_plenum_fasteners_origin = [9481.047852+coefficient, 6939.066406, 8479.819336]
            upper_plenum_spud_fasteners_origin = [9816.714844+coefficient, -1797.713013, 10870.084961]
            all_annotations_sight_direction = [0.585162, 0.594175, -0.551853]
            all_annotations_up_direction = [0.352398, 0.426587, 0.83297]
            engineering_definition_sight_direction = [0.585162, 0.594175, -0.551853]
            engineering_definition_up_direction = [0.352398, 0.426587, 0.83297]
            reference_geometry_sight_direction = [-0.05904, 0.674833, -0.735605]
            reference_geometry_up_direction = [-0.065312, 0.7327, 0.67741]
            upper_plenum_fasteners_sight_direction = [0.067624, -0.955728, -0.286376]
            upper_plenum_fasteners_up_direction = [0.024692, -0.285342, 0.958108]
            upper_plenum_spud_fasteners_sight_direction = [-0.05904, 0.674833, -0.735605]
            upper_plenum_spud_fasteners_up_direction = [-0.064112, 0.732806, 0.677411]

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


class CarmLowerBinNonConstant(CarmLowerBin):

    #def __init__(self, side, *args, **kwargs):
        #super(CARM_UPR, self).__init__(*args, **kwargs)

    def __init__(self, carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name, state, *args):
        super(CarmLowerBin, self).__init__(carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name, *args)
        self.side = side
        self.extention = '\\seed_carm_nonc_lwr_' + side + '.CATPart'

        for i in args:
            if len(i) != 0:
                self.first_elem_in_irm_new = i[0]
        if state == 'final':
            for i in args:
                for n in xrange(7):
                    if '24' in i[n]:
                        continue
                    else:
                        self.first_elem_in_irm = i[n]
                        break
                self.irm_components = i
                self.irm_length = len(i)
                self.first_elem_in_irm_size = self.first_elem_in_irm[:2]

    def set_parameters(self, sta_value_pairs, size):

        carm_part = self.select_carm_to_paste_data()
        #carm_part = self.access_carm()
        parameters1 = carm_part.Parameters
        FL2_X_param = parameters1.Item('FL2_X')
        FL3_X_param = parameters1.Item('FL3_X')
        FL4_X_param = parameters1.Item('FL4_X')
        FL5_X_param = parameters1.Item('FL5_X')
        fir_tree_param = parameters1.Item('156-00066a')
        bacs_param = parameters1.Item('BACS38K2a')

        FL2_X_param_offset = inch_to_mm(4.0606)
        FL3_X_param_offset = inch_to_mm(2.1582)
        FL4_X_param_offset = inch_to_mm(3.0059)
        FL5_X_param_offset = inch_to_mm(2.4261)
        fir_tree_param_offset = inch_to_mm(0.4353)
        bacs_param_offset = inch_to_mm(1.4077)
        anchor_point = self.copy_jd2_bacs12fa3k3_and_paste_1(size, 'first_point')
        k = -1
        anchor_point_coord_X = anchor_point.X
        anchor_point_X = anchor_point_coord_X.Value
        print 'lwrs_params:' + str(anchor_point_X)
        #coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        coord_to_move = -1 * (sta_value_pairs[self.copy_from_product - 5])
        print 'lwrs_params:' + str(coord_to_move)
        print coord_to_move
        print size

        FL2_X_param.Value = coord_to_move + (inch_to_mm(float(size)) - inch_to_mm(0.25) - anchor_point_X) + k * FL2_X_param_offset
        FL3_X_param.Value = coord_to_move + (inch_to_mm(float(size)) - inch_to_mm(0.25) - anchor_point_X) + k * FL3_X_param_offset
        FL4_X_param.Value = coord_to_move + (inch_to_mm(float(size)) - inch_to_mm(0.25) - anchor_point_X) + k * FL4_X_param_offset
        FL5_X_param.Value = coord_to_move + (inch_to_mm(float(size)) - inch_to_mm(0.25) - anchor_point_X) + k * FL5_X_param_offset
        fir_tree_param.Value = coord_to_move + (inch_to_mm(float(size)) - inch_to_mm(0.25) - anchor_point_X) + k * fir_tree_param_offset
        bacs_param.Value = coord_to_move + (inch_to_mm(float(size)) - inch_to_mm(0.25) - anchor_point_X) + k * bacs_param_offset

    def change_inst_id_sta(self, sta_values_fake, side):

        start_sta = sta_values_fake[self.copy_from_product - 5]
        # actual station:
        #finish_sta = int((sta_values_fake[self.copy_from_product - (5 + (self.irm_length - 1))])[1:]) + int(self.first_elem_in_irm_size)
        finish_sta = 465
        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        Product2 = to_p.ReferenceProduct
        instance_id_IRM = 'ECS_LWR-AIR-DIST_INSTL_STA' + start_sta + '-0' + str(finish_sta) + '_' + side[0]
        to_p.Name = instance_id_IRM
        print to_p.Name

    def shift_camera(self, sta_value_pairs, size=None):

        #carm_part = self.select_carm_to_paste_data()
        cam_dict = self.map_camera_names()
        documents1 = self.catia.Documents
        partDocument1 = documents1.Item('CA' + self.carm_part_number + '.CATPart')
        cameras = partDocument1.Cameras

        #coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        #coefficient = coord_to_move - (Inch_to_mm(717.0 - float(size)/2))

        coord_to_move = inch_to_mm(457.917) - (-1 * sta_value_pairs[self.copy_from_product - 5])
        coefficient = coord_to_move - inch_to_mm(457.917)

        print coefficient

        if self.side == 'LH':
            engineering_definition_release_origin = [13944.40332 + coefficient, 2419.441406, 11147.487305]
            all_annotations_origin = [18506.052734 + coefficient, 6726.539063, 15580.635742]
            reference_geometry_origin = [10613.700195 + coefficient, 4517.101563, 7088.336914]
            lower_plenum_downer_strap_origin = [9381.522461 + coefficient, -7979.621094, 7004.487793]
            upper_downer_strap_origin = [9314.950195 + coefficient, -5907.354004, 10929.607422]
            lower_plenum_fastener_jd01_origin = [9201.066406 + coefficient, -8280.686523, 7016.032227]
            lower_plenum_fastener_jd02_origin = [9301.005859 + coefficient, -5827.109375, 10699.782227]
            sidewall_nozzle_fastener_jd03_origin = [10391.533203 + coefficient, 1334.777466, 7020.099121]
            sidewall_nozzle_fastener_jd04_origin = [9310.69043 + coefficient, -8233.225586, 7042.39502]

            engineering_definition_sight_direction = [-0.57735, -0.57735, -0.57735]
            engineering_definition_up_direction = [-0.408248, -0.408248, 0.816497]
            all_annotations_sight_direction = [-0.57735, -0.57735, -0.57735]
            all_annotations_up_direction = [-0.408248, -0.408248, 0.816497]
            reference_geometry_sight_direction = [-0.087156, -0.996195, 0]
            reference_geometry_up_direction = [0, 0, 1]
            lower_plenum_downer_strap_sight_direction = [0.084765, 0.732081, 0.000192]
            lower_plenum_downer_strap_up_direction = [0, 0, 1]
            upper_downer_strap_sight_direction = [0.064049, 0.734878, -0.6782]
            upper_downer_strap_up_direction = [0.059109, 0.675619, 0.734878]
            lower_plenum_fastener_jd01_sight_direction = [0.087156, 0.996195, 0]
            lower_plenum_fastener_jd01_up_direction = [0, 0, 1]
            lower_plenum_fastener_jd02_sight_direction = [0.063359, 0.724197, -0.686676]
            lower_plenum_fastener_jd02_up_direction = [0.059848, 0.684063, 0.726964]
            sidewall_nozzle_fastener_jd03_sight_direction = [-0.087156, -0.996195, 0]
            sidewall_nozzle_fastener_jd03_up_direction = [0, 0, 1]
            sidewall_nozzle_fastener_jd04_sight_direction = [0.087156, 0.996195, 0]
            sidewall_nozzle_fastener_jd04_up_direction = [0, 0, 1]

        else:
            engineering_definition_release_origin = [17127.736328 + coefficient, 8979.547852, 14095.664063]
            all_annotations_origin = [19730.648438 + coefficient, 11429.279297, 16656.181641]
            reference_geometry_origin = [10325.015625 + coefficient, -4908.834961, 7130.697266]
            lower_plenum_downer_strap_origin = [9605.200195 + coefficient, 6803.354004, 6960.905762]
            upper_downer_strap_origin = [10081.00293 + coefficient, -1647.880615, 10959.397461]
            lower_plenum_fastener_jd01_origin = [9499.174805 + coefficient, 8518.854492, 7055.207031]
            lower_plenum_fastener_jd02_origin = [9858.34375 + coefficient, 7175.333008, 11960.477539]
            sidewall_nozzle_fastener_jd03_origin = [10277.820313 + coefficient, -4771.490723, 6982.922363]
            sidewall_nozzle_fastener_jd04_origin = [9663.085938 + coefficient, 8862.500977, 7014.87207]

            engineering_definition_sight_direction = [-0.57735, -0.57735, -0.57735]
            engineering_definition_up_direction = [-0.408248, -0.408248, 0.816497]
            all_annotations_sight_direction = [-0.57735, -0.57735, -0.57735]
            all_annotations_up_direction = [-0.411367, -0.405122, 0.816489]
            reference_geometry_sight_direction = [-0.087156, 0.996195, 0]
            reference_geometry_up_direction = [0, 0, 1]
            lower_plenum_downer_strap_sight_direction = [0.087156, -0.996195, 0]
            lower_plenum_downer_strap_up_direction = [0, 0, 1]
            upper_downer_strap_sight_direction = [-0.059109, 0.675619, -0.734878]
            upper_downer_strap_up_direction = [-0.064049, 0.732081, 0.6782]
            lower_plenum_fastener_jd01_sight_direction = [0.087156, -0.996195, 0]
            lower_plenum_fastener_jd01_up_direction = [0, 0, 1]
            lower_plenum_fastener_jd02_sight_direction = [0.063359, -0.724197, -0.686676]
            lower_plenum_fastener_jd02_up_direction = [0.059848, -0.684063, 0.726964]
            sidewall_nozzle_fastener_jd03_sight_direction = [-0.087156, 0.996195, 0]
            sidewall_nozzle_fastener_jd03_up_direction = [0, 0, 1]
            sidewall_nozzle_fastener_jd04_sight_direction = [0.087156, -0.996195, 0]
            sidewall_nozzle_fastener_jd04_up_direction = [0, 0, 1]

        view_origins = [all_annotations_origin, engineering_definition_release_origin, reference_geometry_origin, lower_plenum_downer_strap_origin, upper_downer_strap_origin, lower_plenum_fastener_jd01_origin, lower_plenum_fastener_jd02_origin, sidewall_nozzle_fastener_jd03_origin, sidewall_nozzle_fastener_jd04_origin]
        viewpoints = ['ALL ANNOTATION', 'ENGINEERING DEFINITION RELEASE', 'REFERENCE GEOMETRY', 'LOWER PLENUM DOWNER STRAP', 'UPPER DOWNER STRAP', 'LOWER PLENUM FASTENER JD01', 'LOWER PLENUM FASTENER JD02', 'SIDEWALL NOZZLE FASTENER JD03', 'SIDEWALL NOZZLE FASTENER JD04']
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

    def add_jd_annotation(self, jd_number, sta_value_pairs, size, side, arch):
        """Adds JOINT DEFINITION XX annotation"""

        annot_text = 'JOINT DEFINITION ' + jd_number
        carm_part = self.access_carm()
        #self.activate_view(jd_number)
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        userSurfaces1 = carm_part.UserSurfaces
        geosets = carm_part.HybridBodies
        geoset1 = geosets.Item('Joint Definitions')
        geosets1 = geoset1.HybridBodies
        geoset2 = geosets1.Item('Joint Definition ' + jd_number)
        points = geoset2.HybridShapes
        if '1' in str(jd_number):
            JD_point = self.copy_jd1_BACS12FA3K20_and_paste(self.first_elem_in_irm_size, 'find_point')
        elif '2' in str(jd_number):
            JD_point = self.copy_jd2_bacs12fa3k3_and_paste_1(self.first_elem_in_irm_size, 'find_point')
        elif '3' in str(jd_number):
            JD_point = self.copy_jd3_BACS12FA3K12_and_paste(self.first_elem_in_irm_size, 'find_point')
        else:
            JD_point = self.copy_jd4_bacs12fa3k3_and_paste_2(self.first_elem_in_irm_size, 'find_point')
        JD_point_coord_X = JD_point.X
        JD_point_X = JD_point_coord_X.Value
        print JD_point_X
        wb = str(self.workbench_id())
        if wb != 'PrtCfg':
            self.swich_to_part_design()
        selection1 = self.productDocument1.Selection
        selection1.Clear()
        reference1 = carm_part.CreateReferenceFromObject(points.Item(1))
        userSurface1 = userSurfaces1.Generate(reference1)

        for point in xrange(2, points.Count+1):
            reference2 = carm_part.CreateReferenceFromObject(points.Item(point))
            print reference2.name
            userSurface1.AddReference(reference2)

        annotationFactory1 = ann_set1.AnnotationFactory
        #coord_to_move = sta_value_pairs[self.copy_from_product - 5 - (self.irm_length - 1)] + JD_point_X


        y = inch_to_mm(12)
        z = 0
        if side == 'LH':
            k = 1
        else:
            k = -1
        if '3' in str(jd_number):
            k = k * (-1)
        addition = k * inch_to_mm(12.0)

        coord_to_move = inch_to_mm(457.917) + sta_value_pairs[self.copy_from_product - 5] - (inch_to_mm(float(size))) + k * JD_point_X
        print coord_to_move

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
        self.activate_view(3)
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        userSurfaces1 = carm_part.UserSurfaces
        geosets = carm_part.HybridBodies
        geoset1 = geosets.Item('Construction Geometry (REF)')
        geosets1 = geoset1.HybridBodies
        geoset2 = geosets1.Item('Misc Construction Geometry')
        points = geoset2.HybridShapes
        hybridShapeFactory1 = carm_part.HybridShapeFactory
        #coord_to_move_ref_point = sta_value_pairs[self.copy_from_product - 5] + Inch_to_mm(0.25)
        coord_to_move_ref_point = -1*(sta_value_pairs[self.copy_from_product - 5]) + (inch_to_mm(float(size))) - inch_to_mm(0.5)
        points = geoset2.HybridShapes
        wb = str(self.workbench_id())
        if wb != 'PrtCfg':
            self.swich_to_part_design()
        reference2 = carm_part.CreateReferenceFromObject(points.Item('point_direction'))
        hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointOnCurveFromDistance(reference2, coord_to_move_ref_point, False)
        geoset2.AppendHybridShape(hybridShapePointCoord1)
        reference1 = carm_part.CreateReferenceFromObject(points.Item(points.Count))
        carm_part.Update()
        userSurface1 = userSurfaces1.Generate(reference1)
        annotationFactory1 = ann_set1.AnnotationFactory
        #coord_to_move = sta_value_pairs[self.copy_from_product - 5] - Inch_to_mm(2)
        coord_to_move = inch_to_mm(457.917) - coord_to_move_ref_point
        y = inch_to_mm(40)
        z = 0
        if side == 'LH':
            k = -1
        else:
            k = 1
        #addition = k*(Inch_to_mm(float(size)/2))*0
        addition = k*(inch_to_mm(2.15))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        text_leaders = ann1text_2d.Leaders
        text_leader1 = text_leaders.Item(1)
        text_leader1.HeadSymbol = 1
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 24)
        print ann1text_2d.AnchorPosition
        ann1text_2d.AnchorPosition = 6
        print ann1text_2d.AnchorPosition
        ann1text_2d.FrameType = 3
        self.rename_part_body()
        self.hide_last_annotation()
        carm_part.Update()

    def add_ref_annotation(self, sta_value_pairs, size, side):
        """Adds REF annotation"""

        annot_text = str(size) + 'IN OUTBD BIN SUPPORT REF'
        carm_part = self.access_carm()
        self.activate_view(3)
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        userSurfaces1 = carm_part.UserSurfaces
        geosets = carm_part.HybridBodies
        geoset1 = geosets.Item('Construction Geometry (REF)')
        geosets1 = geoset1.HybridBodies
        geoset2 = geosets1.Item('Misc Construction Geometry')
        hybridShapeFactory1 = carm_part.HybridShapeFactory
        #coord_to_move_ref_point = sta_value_pairs[self.copy_from_product - 5] + (Inch_to_mm(float(size))*0.7)
        coord_to_move_ref_point = -1*(sta_value_pairs[self.copy_from_product - 5]) + (inch_to_mm(float(size)) * 0.5)
        print 'ref_point: ' + str(coord_to_move_ref_point)
        points = geoset2.HybridShapes
        wb = str(self.workbench_id())
        if wb != 'PrtCfg':
            self.swich_to_part_design()
        reference2 = carm_part.CreateReferenceFromObject(points.Item('point_direction'))
        print reference2.Name
        hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointOnCurveFromDistance(reference2, coord_to_move_ref_point, False)
        geoset2.AppendHybridShape(hybridShapePointCoord1)
        reference1 = carm_part.CreateReferenceFromObject(points.Item(points.Count))
        carm_part.Update()
        r = points.Item(points.Count)
        print r.Name
        userSurface1 = userSurfaces1.Generate(reference1)
        annotationFactory1 = ann_set1.AnnotationFactory
        #coord_to_move = sta_value_pairs[self.copy_from_product - 5]
        coord_to_move = inch_to_mm(457.917) - coord_to_move_ref_point
        y = inch_to_mm(60)
        z = 0
        if side == 'LH':
            k = -1
        else:
            k = 1
        addition = k*(inch_to_mm(float(size) / 2))
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, k * coord_to_move + addition, y, z, True)
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        text_leaders = ann1text_2d.Leaders
        text_leader1 = text_leaders.Item(1)
        text_leader1.HeadSymbol = 20
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 24)
        self.rename_part_body()
        self.hide_last_annotation()
        carm_part.Update()


class CarmUpperBinNonConstantSection47(CarmUpperBin):

    def __init__(self, carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name, *args):
        super(CarmUpperBin, self).__init__(carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name)
        self.side = side
        self.extention = '\\seed_carm_nonc_irm_' + side + '.CATPart'
        for i in args:
            if type(i) is list:
                if len(i) != 0:
                    self.first_elem_in_irm = i[0]
                    self.irm_length = len(i)
                    self.first_elem_in_irm_size = self.first_elem_in_irm[:2]
            else:
                self.plug_value = i

    def change_inst_id_sta(self, sta_values_fake, sta_value_pairs, side, size):

        start_sta = sta_values_fake[self.copy_from_product - (5 + (self.irm_length - 1))]
        finish_sta = int(sta_values_fake[self.copy_from_product - 5]) + int(size)
        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        Product2 = to_p.ReferenceProduct
        #Product2Products = Product2.Products
        #product_forpaste = Product2Products.Item(3)
        instance_id_IRM = 'ECS_UPR-AIR-DIST_INSTL_STA' + start_sta + '-' + str(finish_sta) + '_' + side[0]
        to_p.Name = instance_id_IRM
        print to_p.Name


class CarmLowerBinNonConstantSection47(CarmLowerBin):

    def __init__(self, carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name, state, *args):
        super(CarmLowerBin, self).__init__(carm_part_number, instance_id, side, order_of_new_product, copy_from_product, cfp_name, *args)
        self.side = side
        self.extention = '\\seed_carm_nonc_lwr_' + side + '.CATPart'

        for i in args:
            if type(i) is list:
                if len(i) != 0:
                    self.first_elem_in_irm_new = i[0]
            else:
                self.plug_value = i
        if state == 'final':
            for i in args:
                if type(i) is not list:
                    break
                else:
                    for n in xrange(7):
                        if '24' in i[n]:
                            continue
                        else:
                            self.first_elem_in_irm = i[n]
                            break
                self.irm_length = len(i)
                self.first_elem_in_irm_size = self.first_elem_in_irm[:2]

    def change_inst_id_sta(self, sta_values_fake, sta_value_pairs, side, size):

        start_sta = sta_values_fake[self.copy_from_product - (5 + (self.irm_length - 1))]
        finish_sta = int(sta_values_fake[self.copy_from_product - 5]) + int(size)
        Prod = self.productDocument1.Product
        collection = Prod.Products
        to_p = collection.Item(self.order_of_new_product)
        Product2 = to_p.ReferenceProduct
        #Product2Products = Product2.Products
        #product_forpaste = Product2Products.Item(3)
        instance_id_IRM = 'ECS_LWR-AIR-DIST_INSTL_STA' + start_sta + '-' + str(finish_sta) + '_' + side[0]
        to_p.Name = instance_id_IRM
        print to_p.Name


class PartNumbering(object):

    """
    roll - Z, irm_type - X, nozzle_type - Y
    """

    def __init__(self, start_sta, irm_type, nozzle_type, irm_elements, roll):
        self.start_sta = start_sta
        self.irm_type = irm_type
        self.nozzle_type = nozzle_type
        self.roll = roll
        self.irm_elements = irm_elements

        section1 = 345
        section2 = 465
        section3 = 561
        section4 = 657
        section5 = 897
        section6 = 1089
        section7 = 1293
        section8 = 1401+96
        section9 = 1618
        section10 = 1769

        irm_type_dict = {'UPR':1, 'LWR':2, 'OMF':3}

        nozzle_type_dict = {'economy':4, 'premium':5}

        zone_dict = {xrange(section1, section2):'IR830Z091', xrange(section2, section3):'IR830Z092', xrange(section3, section4):'IR830Z092', xrange(section4, section5):'IR830Z093', xrange(section5, section6):'IR830Z093', xrange(section6, section7):'IR830Z094', xrange(section7, section8):'IR830Z095', xrange(section8, section9):'IR830Z095', xrange(section9, section10):'IR830Z096'}

        config01 = [(24, 42, 48), (48, 48), (48, 48, 48, 36), (24, 36, 48, 48, 48), (48, 48, 48, 48), (48, 48, 48, 36, 24), (36, 48, 48), (48, 48, 48, 48, 48), (48, 36, 24)]
        config03 = [(243, 36, 48), (48, 48), (48, 48, 24, 36, 24), (24, 36, 42, 48, 48), (48, 48, 48, 48), (48, 48, 42, 36, 243), (36, 48, 48), (48, 48, 48, 48, 48), (42, 36, 243)]
        config05 = [(24, 36, 48), (48, 48), (48, 48, 42, 36), (48, 48, 48, 48), (48, 48, 48, 48), (48, 48, 48, 48), (36, 48, 48), (48, 48, 48, 48, 48), (48, 48)]
        config07 = [(48, 48), (48, 48), (48, 42, 24, 36, 24), (24, 24, 48, 48, 48), (48, 48, 48, 48), (48, 48, 48, 24, 24), (36, 48, 48), (48, 48, 48, 48, 48), (48, 24, 24)]
        config09 = [(42, 48), (48, 48), (48, 42, 24, 36, 243), (42, 48, 48, 48), (48, 48, 48, 48), (48, 48, 48, 42), (36, 48, 48), (48, 48, 48, 48, 48), (48, 42)]
        config11 = [(36, 48), (48, 48), (48, 48, 48, 24), (36, 48, 48, 48), (48, 48, 48, 48), (48, 48, 48, 36), (36, 48, 48), (48, 48, 48, 48, 48), (48, 36)]
        config13 = [(36, 42), (48, 48), (48, 36, 36, 42), (36, 42, 48, 48), (48, 48, 48, 48), (48, 48, 42, 36), (36, 48, 48), (48, 48, 48, 48, 48), (42, 36)]
        config15 = [(24, 48), (48, 48), (48, 48, 36, 24), (24, 48, 48, 48), (48, 48, 48, 48), (48, 48, 48, 24), (36, 48, 48), (48, 48, 48, 48, 48), (48, 24)]
        config17 = [(24, 42), (48, 48), (48, 36, 36, 36), (24, 42, 48, 48), (48, 48, 48, 48), (48, 48, 42, 24), (36, 48, 48), (48, 48, 48, 48, 48), (42, 24)]
        config19 = [(24, 36), (48, 48), (48, 42, 24, 36), (24, 36, 48, 48), (48, 48, 48, 48), (48, 48, 36, 24), (36, 48, 48), (48, 48, 48, 48, 48), (36, 24)]
        config21 = [(24, 36), (48, 48), (48, 48, 48), (24, 36, 42, 48), (48, 48, 48, 48), (48, 42, 36, 24), (36, 48, 48), (48, 48, 48, 48, 48), [48]]
        config23 = [[48], (48, 48), (48, 48, 42), (48, 48, 48), (48, 48, 48, 48), (48, 48, 48), (36, 48, 48), (48, 48, 48, 48, 48), [42]]
        config25 = [[48], (48, 48), (48, 48, 36), (42, 48, 48), (48, 48, 48, 48), (48, 48, 42), (36, 48, 48), (48, 48, 48, 48, 48), []]
        config27 = [[42], (48, 48), (48, 42, 36), (42, 48, 48), (48, 48, 48, 48), (48, 48, 42), (36, 48, 48), (48, 48, 48, 48, 48), []]
        config29 = [[42], (48, 48), (48, 48, 24), (36, 48, 48), (48, 48, 48, 48), (48, 48, 36), (36, 48, 48), (48, 48, 48, 48, 48), []]
        config31 = [[36], (48, 48), (48, 42, 24), (36, 48, 48), (48, 48, 48, 48), (48, 48, 36), (36, 48, 48), (48, 48, 48, 48, 48), []]
        config33 = [[36], (48, 48), (48, 42, 24), (36, 42, 48), (48, 48, 48, 48), (48, 42, 36), (36, 48, 48), (48, 48, 48, 48, 48), []]
        config35 = [[], (48, 48), (48, 42, 24), (36, 42, 48), (48, 48, 48, 48), (48, 42, 36), (36, 48, 48), (48, 48, 48, 48, 48), []]
        config37 = [[], (48, 48), (48, 36, 24), (24, 48, 48), (48, 48, 48, 48), (48, 48, 24), (36, 48, 48), (48, 48, 48, 48, 48), []]
        config39 = [[], (48, 48), (48, 36, 24), (24, 48, 48), (48, 48, 48, 48), (48, 48, 24), (36, 48, 48), (48, 48, 48, 48, 48), []]
        config41 = [[], (48, 48), (48, 36, 24), (24, 48, 48), (48, 48, 48, 48), (48, 48, 24), (36, 48, 48), (48, 48, 48, 48, 48), []]

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
