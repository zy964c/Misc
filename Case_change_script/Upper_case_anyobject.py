import win32com.client
catia = win32com.client.Dispatch('catia.application')
productDocument1 = catia.ActiveDocument
documents = catia.Documents
selection1 = productDocument1.Selection
selection1.Clear()
print 'CHOOSE ANNOTATIONS AND CAPTURES'
selection1.SelectElement3(['Annotation', 'AnyObject'], 'CHOOSE ANNOTATIONS', True, 2, False)
for selected in xrange(1, selection1.Count2 + 1):
    selected1 = selection1.Item2(selected)
    selected1_type = selection1.Item2(selected).Type
    if selected1_type == 'Annotation':
        annotation1 = selected1.Value
        if 'Flag Note' in annotation1.Name:
            continue
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        text1 = ann1text_2d.Text
        text1_upper = text1.upper()
        ann1text_2d.Text = text1_upper
        annotation1.ModifyVisu()
    else:
        annotation1 = selected1.Value
        text2 = annotation1.Name
        text2_upper = text2.upper()
        annotation1.Name = text2_upper
        if selected1_type == 'Capture':
            camera1 = annotation1.Camera
            text3 = camera1.Name
            text3_upper = text2_upper
            camera1.Name = text3_upper

raw_input('press ENTER to exit')
