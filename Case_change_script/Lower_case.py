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
selection1.SelectElement3(['Annotation', 'Capture', 'TPSView'],'CHOOSE ANNOTATIONS', True, 2, False)
for selected in xrange(1, selection1.Count2 + 1):
    selected1 = selection1.Item2(selected)
    selected1_type = selection1.Item2(selected).Type
    #print selected1_type
    if selected1_type == 'Annotation':
        annotation1 = selected1.Value
        if 'Flag Note' in annotation1.Name:
            continue
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        text1 = ann1text_2d.Text
        case = 'upper'
        text1_lower = ''
        letters = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890'
        for letter in text1:
            if letter not in letters:
                case = 'upper'
                text1_lower += letter
                continue
            if case == 'upper':
                letter = letter.upper()
                text1_lower += letter
            elif case == 'lower':
                letter = letter.lower()
                text1_lower += letter
            case = 'lower'
        ann1text_2d.Text = text1_lower
        annotation1.ModifyVisu()
    else:
        annotation1 = selected1.Value
        #print annotation1.Name
        text2 = annotation1.Name
        case = 'upper'
        text2_lower = ''
        letters = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890'
        for letter in text2:
            if letter not in letters:
                case = 'upper'
                text2_lower += letter
                continue
            if case == 'upper':
                letter = letter.upper()
                text2_lower += letter
            elif case == 'lower':
                letter = letter.lower()
                text2_lower += letter
            case = 'lower'
        annotation1.Name = text2_lower
        if selected1_type == 'Capture':
            camera1 = annotation1.Camera
            camera1.Name = text2_lower
        #print annotation1.Name
        #annotation1.Current = True
        #annotation1.Current = False
        #annotation1.DisplayCapture()

raw_input('press ENTER to exit')
#print 'DONE!'
