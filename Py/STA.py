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
            
    elif plug_value == 456:
        if coord <= round(Inch_to_mm(609)):
            STA = '0' + str(int(round(coord/25.4)))
        elif coord > round(Inch_to_mm(609)) and coord <= round(Inch_to_mm(609+240)):
            STA = '0609+' + str(int(round(coord/25.4 - 609)))
        elif coord > round(Inch_to_mm(609+240)) and coord <= round(Inch_to_mm(1401 + 240)):
            if (coord/25.4 - 240) < 1000:
              STA = '0' + str(int(round(coord/25.4 - 240)))
            else:
              STA = str(int(round(coord/25.4 - 240)))
        elif coord > round(Inch_to_mm(1401+240)) and coord <= round(Inch_to_mm((1401+240)+120)):
            STA = '1401+' + str(int(round(coord/25.4 - (1401+240))))
        elif coord > round(Inch_to_mm(1401+360)) and coord <= round(Inch_to_mm(1618 + 360)):
            STA = str(int(round(coord/25.4 - 360)))
        elif coord > round(Inch_to_mm(1618+360)) and coord <= round(Inch_to_mm((1618+360)+96)):
            STA = '1618+' + str(int(round(coord/25.4 - (1618+360)))) 
        elif coord > round(Inch_to_mm(1618+360+96)):
            STA = str(int(round(coord/25.4 - (360+96))))
            
    elif plug_value == 0:
        if int(round(coord/25.4)) < 1000:
           STA = '0' + str(int(round(coord/25.4)))
        else:
           STA = str(int(round(coord/25.4)))
           
    return STA

a = raw_input("Please enter STA value: ")
minor_model = raw_input("Please enter minor model: ")
plug_value = 0
if int(minor_model) == 8:
    plug_value = 0
elif int(minor_model) == 9:
    plug_value = 240
elif int(minor_model) == 10:
    plug_value = 456
converted = Inch_to_mm (float(a))
print STAvalue (converted, plug_value)
