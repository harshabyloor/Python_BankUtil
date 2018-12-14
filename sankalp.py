from xlrd import open_workbook

class Arm(object):
    def __init__(self, id, dsp_name, dsp_code, hub_code,ifsc,address):
        #print ("init")
        self.id = id
        self.dsp_name = dsp_name
        self.dsp_code = dsp_code
        self.hub_code = hub_code
        self.ifsc= ifsc
        self.address=address

    def __str__(self):
        #print ("str")
        return("{0}|{1}|{2}|{3}||{4}|{5}||\n"
               .format(self.id, self.dsp_name, self.dsp_code,
                       self.hub_code,self.ifsc,self.address))

class Arm1(object):
    def __init__(self, id, dsp_name, dsp_code, hub_code,ifsc,address):
        #print ("init")
        self.id = id
        self.dsp_name = dsp_name
        self.dsp_code = dsp_code
        self.hub_code = hub_code
        self.ifsc = ifsc
        self.address = address
    def __str__(self):
        #print ("str")
        return ("{0}#{1}#{2}##{3}##{4}#{5}\n"
                .format(self.id, self.dsp_name, self.dsp_code,
                        self.hub_code, self.ifsc, self.address))

class Arm2(object):
    def __init__(self, id, dsp_name, dsp_code, hub_code,ifsc):
        #print ("init")
        self.id = id
        self.dsp_name = dsp_name
        self.dsp_code = dsp_code
        self.hub_code = hub_code
        self.ifsc = ifsc

    def __str__(self):
        #print ("str")
        return("{0}#{1}#{2}##{3}##{4}#\n"
               .format(self.id, self.dsp_name, self.dsp_code,
                       self.hub_code, self.ifsc))



#--------------------------Other bank benificeary-----------------------------
wb = open_workbook('E:/sankalp/OTHER BANK BENEFICIARY REGISTERED.xlsx')
#print("1")
try:
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        #print("2")
        items = []

        rows = []
        for row in range(1, number_of_rows):
            values = []
            for col in range(number_of_columns):
                value  = (sheet.cell(row,col).value)
                try:
                    value = str(int(value))
                    #print("Values: "+value)
                except ValueError:
                    pass
                finally:
                    values.append(value)
            #print ("3")
            item = Arm(*values)
            #print("4")
            items.append(item)

    text_file = open("E:/sankalp/OTHER BANK BENEFICIARY REGISTERED.txt", "w")
    for item in items:
        print (item)
        text_file.write(format(item))
        #print("Accessing one single value (eg. DSPName): {0}".format(item.dsp_name))
        #print("After item")

    text_file.close()
except:
    print("Exception in ")

#--------------------------Other bank transaction-----------------------------
wb = open_workbook('E:/sankalp/OTHER BANK TRANSACTION FILE.xlsx')
#print("1")
try:
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        #print("2")
        items = []

        rows = []
        for row in range(1, number_of_rows):
            values = []
            for col in range(number_of_columns):
                value  = (sheet.cell(row,col).value)
                try:
                    value = str(int(value))
                    #print("Values: "+value)
                except ValueError:
                    pass
                finally:
                    values.append(value)
            #print ("3")
            item = Arm1(*values)
            #print("4")
            items.append(item)

    text_file = open("E:/sankalp/OTHER BANK TRANSACTION FILE.txt", "w")
    for item in items:
        print (item)
        text_file.write(format(item))
        #print("Accessing one single value (eg. DSPName): {0}".format(item.dsp_name))
        #print("After item")

    text_file.close()
except:
    print("Exception in ")

#--------------------------Other bank benificeary-----------------------------
wb = open_workbook('E:/sankalp/SAME BANK TRANSACTION FILE.xlsx')
#print("1")
try:
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        #print("2")
        items = []

        rows = []
        for row in range(1, number_of_rows):
            values = []
            for col in range(number_of_columns):
                value  = (sheet.cell(row,col).value)
                try:
                    value = str(int(value))
                    #print("Values: "+value)
                except ValueError:
                    pass
                finally:
                    values.append(value)
            #print ("3")
            item = Arm2(*values)
            #print("4")
            items.append(item)

    text_file = open("E:/sankalp/SAME BANK TRANSACTION FILE.txt", "w")
    for item in items:
        print (item)
        text_file.write(format(item))
        #print("Accessing one single value (eg. DSPName): {0}".format(item.dsp_name))
        #print("After item")

    text_file.close()
except:
    print("Exception in ")