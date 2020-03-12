# Created by Super Busy Daddy 16/08/2019
# Script 'HelloWorld' displays 'Hello World!" in cell A1 of the 
# current Calc document.
# SCript 'write_my_text' displays the text 'my_text' that we put as argument
# of function 'write_my_text' in Cell A1 and 
# Script 'write_1_to_10' write number 1 to 10 in cells B1 to B10
# of a LibreOffice Calc Document.

# get the doc from the scripting context 
# which is made available to all scripts
desktop = XSCRIPTCONTEXT.getDesktop()
model = desktop.getCurrentComponent()

# access the active sheet
active_sheet = model.CurrentController.ActiveSheet


def HelloWorld(*args):
    """Write 'Hello World!' in Cell A1"""

    # write 'Hello World' in A1
    active_sheet.getCellRangeByName("A1").String = "Hello World!"
    
def write_my_text(my_text):
    """Write what I want in in Cell A1"""
    
    # write in A1
    active_sheet.getCellRangeByName("A1").String = my_text

def write_1_to_10():
    """Write what I want in in Cell A1"""
    for each_number in range(10):
        print("loop at: ", each_number)        
        active_sheet.getCellByPosition(1,each_number).Value = each_number + 1
        
def write_1_to_10_test():
    """Write what I want in in Cell A1"""

    NumberFormats = model.NumberFormats
    locale = model.CharLocale

    for each_number in range(10):
        active_sheet.getCellByPosition(1,each_number).Value = each_number + 1
        cell_before = active_sheet.getCellByPosition(1,each_number)
        active_sheet.getCellByPosition(2,each_number).Value = cell_before.Value  + 10
        # ~ active_sheet.getCellByPosition(3,each_number).String = str(cell_before.NumberFormats)

def check_cell():
    """Loop through all non-empty cells of a column and do somthing with it"""
    
    
def main(*args):
    """Our main program, that call other functions in the order we want"""
    
    write_my_text("My first macro in Python!")
    write_1_to_10()
    print("All done!")
