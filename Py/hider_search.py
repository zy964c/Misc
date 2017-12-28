import win32com.client

catia = win32com.client.Dispatch('catia.application')
productDocument1 = catia.ActiveDocument
#productDocument1.Product.ApplyWorkMode(2)
productSelection1 = catia.ActiveDocument.Selection

products1 = catia.ActiveDocument.Product.Products
print 'Select part to keep in show'
productSelection1.SelectElement3(['AnyObject'],'Select part to keep in show', False, 0, False)
print 'sel-ed'
selected1 = productSelection1.Item2(0).Value
selected1_name = selected1.PartNumber
print selected1_name
productSelection1.Clear
productSelection2 = catia.ActiveDocument.Selection
productSelection2.Search(str('Part Design'.Part.Name != str(selected1_name)))

productSelection2.visProperties.SetShow(1)
productSelection2.Clear
