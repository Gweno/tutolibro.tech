# Created by Gwenole Capp 11/07/2020
# For tutolibro.tech
# Public Domain, feel free to copy, modify, use in your own scripts
# 
# email: gwenole.capp@gmail.com

def context():
    
    # set global variables for context
    
    global desktop
    global model
    global active_sheet
    
    # get the doc from the scripting context 
    # which is made available to all scripts
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    
    # access the active sheet
    active_sheet = model.CurrentController.ActiveSheet
    
def main():
    
    # call function context()
    context()
    
    # get the range of addresses from selection
    oSelection = model.getCurrentSelection()
    oArea = oSelection.getRangeAddress()

    # store the attribute of CellRangeAddress 
    nLeft = oArea.StartColumn
    nTop = oArea.StartRow
    nRight = oArea.EndColumn
    nBottom = oArea.EndRow
    #(note: could the attribute directly instead of using intermediary variable)
    
    # get the Cell Range 
    oRangeSource = active_sheet.getCellRangeByPosition(nLeft, nTop, nRight, nBottom)
    
    # example by name:
    # ~ oRangeSource = active_sheet.getCellRangeByName('A1:C10')
    
    # get data from the Range of cells and store in a tuple
    oDataSource = oRangeSource.getDataArray()
    
    # print to console
    print(oDataSource)
