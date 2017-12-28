import win32com.client
catia = win32com.client.Dispatch('catia.application')
productDocument1 = catia.ActiveDocument
documents = catia.Documents
selection1 = productDocument1.Selection
selection1.Clear()
#singleselection
#selection1.SelectElement2(['Annotation'], 'CHOOSE ANNOTATION', False)
#multipal selection
print 'CHOOSE ANNOTATIONS AND CAPTURES'
selection1.SelectElement3(['Annotation', 'Capture', 'TPSView'], 'CHOOSE ANNOTATIONS', True, 2, False)
for selected in xrange(1, selection1.Count2 + 1):
    selected1 = selection1.Item2(selected)
    selected1_type = selection1.Item2(selected).Type
    #print selected1_type

    if selected1_type == 'Annotation':
        annotation1 = selected1.Value
        #print annotation1.Name
        if 'Flag Note' in annotation1.Name:
            continue
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        text1 = ann1text_2d.Text
        #print text1
        text1_upper = text1.upper()
        #print text1_upper
        ann1text_2d.Text = text1_upper
        annotation1.ModifyVisu()
    else:
        annotation1 = selected1.Value
        #print annotation1.Name
        text2 = annotation1.Name
        text2_upper = text2.upper()
        annotation1.Name = text2_upper
        if selected1_type == 'Capture':
            camera1 = annotation1.Camera
            text3 = camera1.Name
            text3_upper = text2_upper
            camera1.Name = text3_upper
        #print annotation1.Name
        #annotation1.Current = True
        #annotation1.Current = False
        #annotation1.DisplayCapture()

raw_input('press ENTER to exit')
#print 'DONE!'
