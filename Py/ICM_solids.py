import Tkinter, Tkconstants, tkFileDialog, tkMessageBox
#from Tkinter import IntVar
import math
import win32com.client

a = win32com.client.Dispatch('catia.application')
ICM = a.Documents.Add("Product")
oFileSys = a.FileSystem

ICM_1 = ICM.Product
ICM_1.PartNumber = "ICM"
ICM_1.Name = "ICM"
ICM_Products = ICM_1.Products

new_component1 = ICM_Products.AddNewProduct('non-constant_41_LH')
ICM_Sec41_LH_Products = new_component1.Products

new_component2 = ICM_Products.AddNewProduct('non-constant_41_RH')
ICM_Sec41_RH_Products = new_component2.Products

new_component3 = ICM_Products.AddNewProduct('non-constant_47_LH')
ICM_Sec47_LH_Products = new_component3.Products

new_component4 = ICM_Products.AddNewProduct('non-constant_47_RH')
ICM_Sec47_RH_Products = new_component4.Products


class TkFileDialogExample(Tkinter.Frame):

  def __init__(self, root):

    Tkinter.Frame.__init__(self, root)
    
    #f = StringVar()
    self.plug = Tkinter.IntVar()
    root.geometry("350x180")

    # options for buttons
    button_opt = {'fill': Tkconstants.BOTH, 'padx': 5, 'pady': 5}

    # define buttons
    Tkinter.Button(self, text='Make ICM', command=self.askopenfile).pack(**button_opt)
    Tkinter.Radiobutton(root, text="-8", variable=self.plug, value=0).pack(**button_opt)
    Tkinter.Radiobutton(root, text="-9", variable=self.plug, value=240).pack(**button_opt)
    Tkinter.Radiobutton(root, text="-10", variable=self.plug, value=456).pack(**button_opt)
    
    self.plug.set(240)
    

    # define options for opening or saving a file
    self.file_opt = options = {}
    options['defaultextension'] = '.txt'
    options['filetypes'] = [('all files', '.*'), ('text files', '.txt')]
    options['initialdir'] = 'C:\\'
    options['initialfile'] = 'myfile.txt'
    options['parent'] = root
    options['title'] = 'This is a title'

    # This is only available on the Macintosh, and only when Navigation Services are installed.
    #options['message'] = 'message'

    # if you use the multiple file version of the module functions this option is set automatically.
    #options['multiple'] = 1
   
  def askopenfile(self):

    """Returns an opened file in read mode."""
    
    f = tkFileDialog.askopenfile(parent=root,mode='rb',title='Choose a file')

    s1 = f.readline()
    if s1.startswith("#"):
     s1 = f.readline().replace(' ', '').replace('fairing', '1').replace('prem', '2').replace('EXT', '3').split(",")
    s2 = f.readline()
    if s2.startswith("#"):
     s2 = f.readline().replace(' ', '').replace('fairing', '1').replace('prem', '2').replace('EXT', '3').split(",")
    s3 = f.readline()
    if s3.startswith("#"):
     s3 = f.readline().replace(' ', '').replace('fairing', '1').replace('prem', '2').replace('EXT', '3').split(",")  
    s4 = f.readline()
    if s4.startswith("#"):
     s4 = f.readline().replace(' ', '').replace('fairing', '1').replace('prem', '2').replace('EXT', '3').split(",")   
    s5 = f.readline()
    if s5.startswith("#"):
     s5 = f.readline().replace(' ', '').replace('fairing', '1').replace('prem', '2').replace('EXT', '3').split(",")
    s6 = f.readline()
    if s6.startswith("#"):
     s6 = f.readline().replace(' ', '').replace('fairing', '1').replace('prem', '2').replace('EXT', '3').split(",")
 
    s1 = s1[::-1]
    s2 = s2[::-1]

    print s1 # just for checking list correctness
    print s2 # just for checking list correctness
    print s3 # just for checking list correctness
    print s4 # just for checking list correctness
    print s5 # just for checking list correctness
    print s6 # just for checking list correctness
    
    if self.plug.get() == 0:
        plug = 0
    elif self.plug.get() == 456:
        plug = 456
    else:
        plug = 240

 
    AddComponent(s3, 'LH', 'constant', 'middle', plug)
    AddComponent(s4, 'RH', 'constant', 'middle', plug)
    AddComponent(s1, 'LH', 'nonconstant', 'nose', 0)
    AddComponent(s2, 'RH', 'nonconstant', 'nose', 0)
    AddComponent(s5, 'LH', 'nonconstant', 'tail', plug)
    AddComponent(s6, 'RH', 'nonconstant', 'tail', plug)
    
    
    tkMessageBox.showinfo(title="ICM", message="Done")
    root.withdraw()
    root.destroy()
    
    
#f = tkFileDialog.askopenfile(parent=root,mode='rb',title='Choose a file')
#root.withdraw()
#f = open('C:\Documents and Settings\Roman\My documents\CATIA_V5\Test\ANA.txt', 'r')


angle = 0

def Inch_to_mm (distance):
    return distance * 25.4
    
def STAvalue (coord, plug_value):
    if plug_value == 240:
        if coord <= round(Inch_to_mm(609)):
            STA = '0' + str(int(round(coord/25.4)))
        elif coord > round(Inch_to_mm(609)) and coord <= round(Inch_to_mm(609+120)):
            STA = '0609+' + str(int(round(coord/25.4 - 609)))
        elif coord > round(Inch_to_mm(609+120)) and coord <= round(Inch_to_mm(1401 + 120)):
            if (coord/25.4 - 120) < 1000:
              STA = '0' + str(int(round(coord/25.4 - 120)))
            else:
              STA = str(int(round(coord/25.4 - 120)))
        elif coord > round(Inch_to_mm(1401+120)) and coord <= round(Inch_to_mm((1401+120)+120)):
            STA = '1401+' + str(int(round(coord/25.4 - (1401+120))))
        elif coord > round(Inch_to_mm(1401+240)):
            STA = str(int(round(coord/25.4 - 240)))
            
    elif plug_value == 0:
        STA = '0' + str(int(round(coord/25.4)))
    return STA
        
    



def AddComponent(s, side, section, location, plug_value):
    
 path = 'C:\Documents and Settings\Roman\My documents\CATIA_V5\ICM_SOLIDS'
 
 extention = '.CATProduct'
    
 x_coord = Inch_to_mm(465)
 x_coord_nonconstant = Inch_to_mm(0)
 fake_coord_nonconstant_41 = Inch_to_mm(459)
 fake_coord_nonconstant_47 = Inch_to_mm(1863)
 
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
       
       if str(number) == 'door':
           x_coord = Inch_to_mm(693 + door2_coord)
           index += 1
           continue
           
       else:
           
          Rotate5 = [0.996194698, -0.087155742, 0, 0.087155742, 0.996194698, 0, 0, 0, 1, Inch_to_mm (466.61647022), Inch_to_mm (0.08471639), 0]
          Rotate185 = [-0.996194698, -0.087155742, 0, 0.087155742, -0.996194698, 0, 0, 0, 1, Inch_to_mm (466.61647018), Inch_to_mm (-0.084716377), 0]
          Rotate_5 = [0.998512978, 0.054514501, 0, -0.054514501, 0.998512978, 0, 0, 0, 1, Inch_to_mm (1618.61663822 + plug_value), Inch_to_mm (0.17865996), 0]
          Rotate_185 = [-0.998512978, 0.054514501, 0, -0.054514501, -0.998512978, 0, 0, 0, 1, Inch_to_mm (1618.61663822 + plug_value), Inch_to_mm (-0.17865996), 0]

   
              
          print int(number)# check
          
          # checking area around DOOR 2:
          
          if index != (len(s) - 1) and (s[index + 1] == 'door' or s[index - 1] == 'door'):
            
            if (side == 'LH' and s[index + 1] == 'door') or (side == 'RH' and s[index - 1] == 'door'):
                
              print 'RH door2'
              
              if int(number) == 24:
                iteration += 1
                index += 1
                PartDocPath = path + '\Twenty_four_arch_RH_solids'
                PartDocPath1 = PartDocPath + str(iteration) + extention
                oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                PartDoc = a.Documents.Open(PartDocPath1)
                
              elif int(number) == 243:
                number = '30'
                iteration += 1
                index += 1
                PartDocPath = path + '\Twenty_four_arch_EXT_RH_solids'
                PartDocPath1 = PartDocPath + str(iteration) + extention
                oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                PartDoc = a.Documents.Open(PartDocPath1)
                
              elif int(number) == 18:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Eighteen_arch_RH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 30:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Thirty_arch_RH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 54:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Fifty_four_arch_RH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 60:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Sixty_arch_RH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 241:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Twenty_four_fairing_arch_RH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '24'
            
              elif int(number) == 361:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Thirty_six_fairing_arch_RH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '36'
            
              elif int(number) == 421:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Fourty_two_fairing_arch_RH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '42'
            
              elif int(number) == 481:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Fourty_eight_fairing_arch_RH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '48'
                 
              #PREMIUM:
              
              elif int(number) == 242:
                number = '24'
                nozzl_type = 'PREM' 
                iteration += 1
                index += 1
                PartDocPath = path + '\Twenty_four_arch_RH_solids_pr'
                PartDocPath1 = PartDocPath + str(iteration) + extention
                oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                PartDoc = a.Documents.Open(PartDocPath1)
                
              elif int(number) == 2432:
                number = '30'
                nozzl_type = 'PREM'
                iteration += 1
                index += 1
                PartDocPath = path + '\Twenty_four_arch_EXT_RH_solids_pr'
                PartDocPath1 = PartDocPath + str(iteration) + extention
                oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                PartDoc = a.Documents.Open(PartDocPath1)
                
              elif int(number) == 182:
                 number = '18'
                 nozzl_type = 'PREM'
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Eighteen_arch_RH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 302:
                 number = '30'
                 nozzl_type = 'PREM' 
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Thirty_arch_RH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 542:
                 number = '54'
                 nozzl_type = 'PREM'
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Fifty_four_arch_RH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 602:
                 number = '60'
                 nozzl_type = 'PREM'
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Sixty_arch_RH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 2412:
                 number = '24'
                 nozzl_type = 'PREM'
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Twenty_four_fairing_arch_RH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '24'
            
              elif int(number) == 3612:
                 number = '36'
                 nozzl_type = 'PREM'
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Thirty_six_fairing_arch_RH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '36'
            
              elif int(number) == 4212:
                 number = '42'
                 nozzl_type = 'PREM'
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Fourty_two_fairing_arch_RH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '42'
            
              elif int(number) == 4812:
                 number = '48'
                 nozzl_type = 'PREM'
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Fourty_eight_fairing_arch_RH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '48'
                 
              else:
                  x_coord += Inch_to_mm(int(number))
                  iteration += 1
                  index += 1
                  continue
                 
            elif (side == 'LH' and s[index - 1] == 'door') or (side == 'RH' and s[index + 1] == 'door'):
                
              print 'LH door2'
                
              if int(number) == 24:
                iteration += 1
                index += 1
                PartDocPath = path + '\Twenty_four_arch_LH_solids'
                PartDocPath1 = PartDocPath + str(iteration) + extention
                oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                PartDoc = a.Documents.Open(PartDocPath1)
                
              elif int(number) == 243:
                number = '30'
                iteration += 1
                index += 1
                PartDocPath = path + '\Twenty_four_arch_EXT_LH_solids'
                PartDocPath1 = PartDocPath + str(iteration) + extention
                oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                PartDoc = a.Documents.Open(PartDocPath1)
                
              elif int(number) == 18:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Eighteen_arch_LH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 30:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Thirty_arch_LH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 54:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Fifty_four_arch_LH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 60:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Sixty_arch_LH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 241:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Twenty_four_fairing_arch_LH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '24'
            
              elif int(number) == 361:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Thirty_six_fairing_arch_LH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '36'
            
              elif int(number) == 421:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Fourty_two_fairing_arch_LH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '42'
            
              elif int(number) == 481:
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Fourty_eight_fairing_arch_LH_solids'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '48'
                 
                 #PREM:
                 
              elif int(number) == 242:
                number = '24'
                nozzl_type = 'PREM' 
                iteration += 1
                index += 1
                PartDocPath = path + '\Twenty_four_arch_LH_solids_pr'
                PartDocPath1 = PartDocPath + str(iteration) + extention
                oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                PartDoc = a.Documents.Open(PartDocPath1)
                
              elif int(number) == 2432:
                number = '30'
                nozzl_type = 'PREM'
                iteration += 1
                index += 1
                PartDocPath = path + '\Twenty_four_arch_EXT_LH_solids_pr'
                PartDocPath1 = PartDocPath + str(iteration) + extention
                oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                PartDoc = a.Documents.Open(PartDocPath1)
                
              elif int(number) == 182:
                 number = '18'
                 nozzl_type = 'PREM'
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Eighteen_arch_LH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 302:
                 number = '30'
                 nozzl_type = 'PREM' 
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Thirty_arch_LH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 542:
                 number = '54'
                 nozzl_type = 'PREM'
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Fifty_four_arch_LH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 602:
                 number = '60'
                 nozzl_type = 'PREM'
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Sixty_arch_LH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
              
              elif int(number) == 2412:
                 number = '24'
                 nozzl_type = 'PREM'
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Twenty_four_fairing_arch_LH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '24'
            
              elif int(number) == 3612:
                 number = '36'
                 nozzl_type = 'PREM'
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Thirty_six_fairing_arch_LH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '36'
            
              elif int(number) == 4212:
                 number = '42'
                 nozzl_type = 'PREM'
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Fourty_two_fairing_arch_LH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '42'
            
              elif int(number) == 4812:
                 number = '48'
                 nozzl_type = 'PREM'
                 iteration += 1
                 index += 1
                 PartDocPath = path + '\Fourty_eight_fairing_arch_LH_solids_pr'
                 PartDocPath1 = PartDocPath + str(iteration) + extention
                 oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                 PartDoc = a.Documents.Open(PartDocPath1)
                 number = '48'
                 
              else:
                  x_coord += Inch_to_mm(int(number))
                  iteration += 1
                  index += 1
                  continue
                 
           #  NOT around DOOR 2:
             
          elif location == 'nose' or location == 'middle':
          
           if int(number) == 24:
               
              iteration += 1
              index += 1
              PartDocPath = path + '\Twenty_four_solids'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           elif int(number) == 243:
                number = '30'
                iteration += 1
                index += 1
                if side == 'RH':
                  PartDocPath = path + '\Twenty_four_EXT_DR3_RH_solids'
                else:
                  PartDocPath = path + '\Twenty_four_EXT_DR3_LH_solids'
                PartDocPath1 = PartDocPath + str(iteration) + extention
                oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                PartDoc = a.Documents.Open(PartDocPath1)
             
           elif int(number) == 36:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Thirty_six_solids'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
             
           elif int(number) == 42:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Fourty_two_solids'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
             
           elif int(number) == 48:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Fourty_eight_solids'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
             
           elif int(number) == 12:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Twelve_solids'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           elif int(number) == 18:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Eighteen_solids'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           elif int(number) == 30:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Thirty_solids'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           elif int(number) == 54:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Fifty_four_solids'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           elif int(number) == 60:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Sixty_solids'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           elif int(number) == 72:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Seventy_two_solids'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           elif int(number) == 241:
              
              number = '24'
              iteration += 1
              index += 1
              PartDocPath = path + '\Twenty_four_fairing_solids'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
            
           elif int(number) == 361:
              
              number = '36'
              iteration += 1
              index += 1
              PartDocPath = path + '\Thirty_six_fairing_solids'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
            
           elif int(number) == 421:
              
              number = '42'
              iteration += 1
              index += 1
              PartDocPath = path + '\Fourty_two_fairing_solids'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
            
           elif int(number) == 481:
              
              number = '48'
              iteration += 1
              index += 1
              PartDocPath = path + '\Fourty_eight_fairing_solids'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
              #Premium plenums:
              
           elif int(number) == 242:
               
              number = '24'
              nozzl_type = 'PREM' 
              iteration += 1
              index += 1
              PartDocPath = path + '\Twenty_four_solids_pr'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           elif int(number) == 2432:
                number = '30'
                nozzl_type == 'PREM'
                iteration += 1
                index += 1
                if side == 'RH':
                  PartDocPath = path + '\Twenty_four_EXT_DR3_RH_solids_pr'
                else:
                  PartDocPath = path + '\Twenty_four_EXT_DR3_LH_solids_pr'
                PartDocPath1 = PartDocPath + str(iteration) + extention
                oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                PartDoc = a.Documents.Open(PartDocPath1)
             
           elif int(number) == 362:
               
              number = '36'
              nozzl_type = 'PREM'
              iteration += 1
              index += 1
              PartDocPath = path + '\Thirty_six_solids_pr'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
             
           elif int(number) == 422:
               
              number = '42'
              nozzl_type = 'PREM'
              iteration += 1
              index += 1
              PartDocPath = path + '\Fourty_two_solids_pr'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
             
           elif int(number) == 482:
               
              number = '48'
              nozzl_type = 'PREM'
              iteration += 1
              index += 1
              PartDocPath = path + '\Fourty_eight_solids_pr'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
             
           elif int(number) == 122:
               
              number = '12'
              nozzl_type = 'PREM'
              iteration += 1
              index += 1
              PartDocPath = path + '\Twelve_solids_pr'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           elif int(number) == 182:
              
              number = '18'
              nozzl_type = 'PREM'
              iteration += 1
              index += 1
              PartDocPath = path + '\Eighteen_solids_pr'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           elif int(number) == 302:
               
              number = '30'
              nozzl_type = 'PREM'
              iteration += 1
              index += 1
              PartDocPath = path + '\Thirty_solids_pr'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           elif int(number) == 542:
              
              number = '54'
              nozzl_type = 'PREM'
              iteration += 1
              index += 1
              PartDocPath = path + '\Fifty_four_solids_pr'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           elif int(number) == 602:
              
              number = '60'
              nozzl_type = 'PREM'
              iteration += 1
              index += 1
              PartDocPath = path + '\Sixty_solids_pr'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           elif int(number) == 722:
              
              number = '72'
              nozzl_type = 'PREM'
              iteration += 1
              index += 1
              PartDocPath = path + '\Seventy_two_solids_pr'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           elif int(number) == 2412:
              
              number = '24'
              nozzl_type = 'PREM'
              iteration += 1
              index += 1
              PartDocPath = path + '\Twenty_four_fairing_solids_pr'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
            
           elif int(number) == 3612:
              
              number = '36'
              nozzl_type = 'PREM'
              iteration += 1
              index += 1
              PartDocPath = path + '\Thirty_six_fairing_solids_pr'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
            
           elif int(number) == 4212:
              
              number = '42'
              nozzl_type = 'PREM'
              iteration += 1
              index += 1
              PartDocPath = path + '\Fourty_two_fairing_solids_pr'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
            
           elif int(number) == 4812:
              
              number = '48'
              nozzl_type = 'PREM'
              iteration += 1
              index += 1
              PartDocPath = path + '\Fourty_eight_fairing_solids_pr'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
           else:
                  x_coord += Inch_to_mm(int(number))
                  iteration += 1
                  index += 1
                  continue
                  
              #section 47: 
                 
          else:
              
            if int(number) == 24:
              iteration += 1
              index += 1
              PartDocPath = path + '\Twenty_four_solids_sec47'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
             
            elif int(number) == 36:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Thirty_six_solids_sec47'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
             
            elif int(number) == 42:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Fourty_two_solids_sec47'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
             
            elif int(number) == 48:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Fourty_eight_solids_sec47'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
             
            elif int(number) == 12:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Twelve_solids_sec47'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
            elif int(number) == 18:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Eighteen_solids_sec47'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
            elif int(number) == 30:
              
              iteration += 1
              index += 1
              PartDocPath = path + '\Thirty_solids_sec47'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
            elif int(number) == 241:
              
              number = '24'
              iteration += 1
              index += 1
              PartDocPath = path + '\Twenty_four_fairing_solids_sec47'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
            
            elif int(number) == 361:
              
              number = '36'
              iteration += 1
              index += 1
              PartDocPath = path + '\Thirty_six_fairing_solids_sec47'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
            
            elif int(number) == 421:
              
              number = '42'
              iteration += 1
              index += 1
              PartDocPath = path + '\Fourty_two_fairing_solids_sec47'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
            
            elif int(number) == 481:
              
              number = '48'
              iteration += 1
              index += 1
              PartDocPath = path + '\Fourty_eight_fairing_solids_sec47'
              PartDocPath1 = PartDocPath + str(iteration) + extention
              oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
              PartDoc = a.Documents.Open(PartDocPath1)
              
            else:
                  x_coord += Inch_to_mm(int(number))
                  iteration += 1
                  index += 1
                  continue  
                  
                  
          if section == 'constant':
              
              NewComponent = ICM_Products.AddExternalComponent(PartDoc)
              PartDoc.Close()
              oFileSys.DeleteFile(PartDocPath1)
              NewComponent.PartNumber = str(number) + '_'+ nozzl_type + '_' + str(iteration)
              NewComponent.Name = str(number) + 'IN STA ' + STAvalue(x_coord, plug_value) + ' ' + side
              RenamingTool = NewComponent.ReferenceProduct
              PlenumAssy = RenamingTool.Products.Item(1)
              LING_VAL = RenamingTool.Products.Item(2)
              PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + STAvalue(x_coord, plug_value) + '_' + side[0]
              LING_VAL.name = 'OB_BIN_LIGVAL_STA' + STAvalue(x_coord, plug_value) + '_' + side[0]
              
          elif section == 'nonconstant' and side == 'LH'and location == 'nose':
              
              NewComponent = ICM_Sec41_LH_Products.AddExternalComponent(PartDoc)
              PartDoc.Close()
              oFileSys.DeleteFile(PartDocPath1)
              RenamingToolProd = new_component1.ReferenceProduct
              Prod = RenamingToolProd.Products.Item(index)
              Prod.PartNumber = str(number) + '_'+ nozzl_type + '_' + str(iteration)
              Prod.Name = str(number) + 'IN STA ' + STAvalue((fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + ' ' + side
              RenamingTool = NewComponent.ReferenceProduct
              PlenumAssy = RenamingTool.Products.Item(1)
              LING_VAL = RenamingTool.Products.Item(2)
              PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + STAvalue((fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + side[0]
              LING_VAL.name = 'OB_BIN_LIGVAL_STA' + STAvalue((fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + side[0]
              NewComponent.Move.Apply(Rotate5)
              
          elif section == 'nonconstant' and side == 'RH' and location == 'nose':
              
              NewComponent = ICM_Sec41_RH_Products.AddExternalComponent(PartDoc)
              PartDoc.Close()
              oFileSys.DeleteFile(PartDocPath1)
              RenamingToolProd = new_component2.ReferenceProduct
              Prod = RenamingToolProd.Products.Item(index)
              Prod.PartNumber = str(number) + '_'+ nozzl_type + '_' + str(iteration)
              Prod.Name = str(number) + 'IN STA ' + STAvalue((fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + ' ' + side
              RenamingTool = NewComponent.ReferenceProduct
              PlenumAssy = RenamingTool.Products.Item(1)
              LING_VAL = RenamingTool.Products.Item(2)
              PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + STAvalue((fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + side[0]
              LING_VAL.name = 'OB_BIN_LIGVAL_STA' + STAvalue((fake_coord_nonconstant_41 + x_coord_nonconstant - Inch_to_mm(int(number))), plug_value) + '_' + side[0]
              NewComponent.Move.Apply(Rotate185)
              
          elif section == 'nonconstant' and side == 'LH' and location == 'tail':
              
              NewComponent = ICM_Sec47_LH_Products.AddExternalComponent(PartDoc)
              PartDoc.Close()
              oFileSys.DeleteFile(PartDocPath1)
              RenamingToolProd = new_component3.ReferenceProduct
              Prod = RenamingToolProd.Products.Item(index)
              Prod.PartNumber = str(number) + '_'+ nozzl_type + '_' + str(iteration)
              Prod.Name = str(number) + 'IN STA ' + STAvalue((fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + ' ' + side
              RenamingTool = NewComponent.ReferenceProduct
              PlenumAssy = RenamingTool.Products.Item(1)
              LING_VAL = RenamingTool.Products.Item(2)
              PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + STAvalue((fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
              LING_VAL.name = 'OB_BIN_LIGVAL_STA' + STAvalue((fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
              NewComponent.Move.Apply(Rotate_5)
              
          elif section == 'nonconstant' and side == 'RH' and location == 'tail':
              
              NewComponent = ICM_Sec47_RH_Products.AddExternalComponent(PartDoc)
              PartDoc.Close()
              oFileSys.DeleteFile(PartDocPath1)
              RenamingToolProd = new_component4.ReferenceProduct
              Prod = RenamingToolProd.Products.Item(index)
              Prod.PartNumber = str(number) + '_'+ nozzl_type + '_' + str(iteration)
              Prod.Name = str(number) + 'IN STA ' + STAvalue((fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + ' ' + side
              RenamingTool = NewComponent.ReferenceProduct
              PlenumAssy = RenamingTool.Products.Item(1)
              LING_VAL = RenamingTool.Products.Item(2)
              PlenumAssy.name = str(number) + nozzl_type + 'NOZASSY_' + 'STA' + STAvalue((fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
              LING_VAL.name = 'OB_BIN_LIGVAL_STA' + STAvalue((fake_coord_nonconstant_47 + x_coord_nonconstant), plug_value) + '_' + side[0]
              NewComponent.Move.Apply(Rotate_185)
              
              
          if location == 'nose':
              x_coord_nonconstant -= Inch_to_mm(int(number))
           
          x = x_coord_nonconstant * math.cos(rad)
          y = x_coord_nonconstant * math.sin(rad)
          
          position = [1, 0, 0, 0, 1, 0, 0, 0, 1, x_coord, 0, 0]
          position_non = [1, 0, 0, 0, 1, 0, 0, 0, 1, x, -y, 0]
          position_non_RH = [1, 0, 0, 0, 1, 0, 0, 0, 1, x+(Inch_to_mm(int(number))*math.cos(rad)), y+(Inch_to_mm(int(number))*math.sin(rad)), 0]
          position90 = [-1, 0, 0, 0, -1, 0, 0, 0, 1, x_coord+Inch_to_mm(int(number)), 0, 0] # 90 deg rotation
          position_non_47 = [1, 0, 0, 0, 1, 0, 0, 0, 1, x, y, 0]
          position_non_47_RH = [1, 0, 0, 0, 1, 0, 0, 0, 1, x+(Inch_to_mm(int(number))*math.cos(rad)), (y+(Inch_to_mm(int(number))*math.sin(rad)))*(-1), 0]
            
                 
          if  side == 'LH'and section == 'constant':
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
          elif section == 'nonconstant' and side == 'RH'and location == 'nose':
                 NewComponent.Move.Apply(position_non_RH)
                 print section
                 print x_coord_nonconstant  
          elif section == 'nonconstant' and side == 'LH'and location == 'tail':
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
           
    #print iteration # check
    #print x_coord # check
           

if __name__=='__main__':
  root = Tkinter.Tk()
  TkFileDialogExample(root).pack()
  root.mainloop()