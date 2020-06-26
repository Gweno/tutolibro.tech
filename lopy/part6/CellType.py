# Created by Gwenole Capp 18/06/2020
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
        
    # get the first cell
    firstRow = oArea.StartRow
    firstCol = oArea.StartColumn
    selectedCell = active_sheet.getCellByPosition(firstCol,firstRow)
    
    # get the type of the cell
    cellType = selectedCell.Type.value
    
    # print in console
    print("Cell (",firstCol,",",firstRow,") Type:", cellType)
    
    # display in next cell
    active_sheet.getCellByPosition(firstCol+1,firstRow).String = "Cell Type:" + cellType
