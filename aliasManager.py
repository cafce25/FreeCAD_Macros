# -*- coding: utf-8 -*-
''' Script for setting, moving and clearing aliases in the Spreadsheet.
    It allows to generate Part Families

    hatari 2016 v0.2

    GNU Lesser General Public License (LGPL)
'''

from PySide import QtGui, QtCore
import os
from enum import Enum, auto
import FreeCAD

class Option( Enum):
    """Options to select from"""
    SET_ALIASES = auto()
    CLEAR_ALIASES = auto()
    MOVE_ALIASES = auto()
    GENERATE_PART_FAMILY = auto()

class MyButtons(QtGui.QDialog):
    """"""
    def __init__(self):
        super(MyButtons, self).__init__()
        self.initUI()

    def initUI(self):
        setAliasesButton = QtGui.QPushButton("Set Aliases")
        setAliasesButton.clicked.connect(self.onSetAliases)
        clearAliasesButton = QtGui.QPushButton("Clear Aliases")
        clearAliasesButton.clicked.connect(self.onClearAliases)
        moveAliasesButton = QtGui.QPushButton("Move Aliases")
        moveAliasesButton.clicked.connect(self.onMoveAliases)
        generatePartFamilyButton = QtGui.QPushButton("Generate Part Family")
        generatePartFamilyButton.clicked.connect(self.onGeneratePartFamily)

        buttonBox = QtGui.QDialogButtonBox(QtCore.Qt.Vertical)

        buttonBox.addButton(setAliasesButton, QtGui.QDialogButtonBox.ActionRole)
        buttonBox.addButton(clearAliasesButton, QtGui.QDialogButtonBox.ActionRole)
        buttonBox.addButton(moveAliasesButton, QtGui.QDialogButtonBox.ActionRole)
        buttonBox.addButton(generatePartFamilyButton, QtGui.QDialogButtonBox.ActionRole)
        
        mainLayout = QtGui.QVBoxLayout()
        mainLayout.addWidget(buttonBox)
        self.setLayout(mainLayout)
        
        # define window		xLoc,yLoc,xDim,yDim
        # self.setGeometry(400, 400, 300, 50)
        self.setWindowTitle("Alias Manager")
        self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)

    def onSetAliases(self):
        self.retStatus = Option.SET_ALIASES
        self.close()
    def onClearAliases(self):
        self.retStatus = Option.CLEAR_ALIASES
        self.close()
    def onMoveAliases(self):
        self.retStatus = Option.MOVE_ALIASES
        self.close()
    def onGeneratePartFamily(self):
        self.retStatus = Option.GENERATE_PART_FAMILY
        self.close()

def setAliases():
    doc = App.ActiveDocument
    sheet = doc.Spreadsheet

    column = QtGui.QInputDialog.getText(None, "Column containing Values", "Enter Column Letter")
    if not column[1]:
        return
    col = column[0].upper() # Always use capital characters for Spreadsheet
    startCell =  QtGui.QInputDialog.getInt(None, "start Row number", "Input start Row number:")
    if not startCell[1]:
        return
    endCell =  QtGui.QInputDialog.getInt(None, "end Row number", "Input end Row number:")
    if not endCell[1]:
        return

    for i in range(startCell[0],endCell[0]+1):
        cellFrom = 'A' + str(i)
        cellTo = col + str(i)
        sheet.setAlias(cellTo, '')
        doc.recompute()
        sheet.setAlias(cellTo, sheet.getContents(cellFrom))


def clearAliases():
    doc = FreeCAD.ActiveDocument
    column = QtGui.QInputDialog.getText(None, "Column containing Values", "Enter Column Letter")
    if not column[1]:
        return
    col = str.capitalize(str(column[0]))
    startCell =  QtGui.QInputDialog.getInt(None, "start Row number", "Input start Row number:")
    if not startCell[1]:
        return
    endCell =  QtGui.QInputDialog.getInt(None, "end Row number", "Input end Row number:")
    if not endCell[1]:
        return
    for i in range(startCell[0],endCell[0]+1):
        cellTo = str(col[0]) + str(i)
        doc.Spreadsheet.setAlias(cellTo, '')
        doc.recompute()

def moveAliases():
    doc = FreeCAD.ActiveDocument
    sheet = doc.Spreadsheet

    columnFrom = QtGui.QInputDialog.getText(None, "Value Column", "Move From")
    if not columnFrom[1]:
        return
    columnTo = QtGui.QInputDialog.getText(None, "Value Column", "Move To")
    if not columnTo[1]:
        return
    columnFrom = columnFrom[0].upper()
    columnTo = columnTo[0].upper()
    startCell =  QtGui.QInputDialog.getInt(None, "start Row number", "Input start Row number:")
    if not startCell[1]:
        return
    endCell =  QtGui.QInputDialog.getInt(None, "end Row number", "Input end Row number:")
    if not endCell[1]:
        return

    for i in range(startCell[0],endCell[0]+1):
        cellDef = 'A'+ str(i)                        
        cellFrom = columnFrom + str(i)
        cellTo = columnTo + str(i)
        sheet.setAlias(cellFrom, '')
        doc.recompute()
        sheet.setAlias(cellTo, sheet.getContents(cellDef))

def charRange( a, b):
    
    """Generates the characters from `c1` to `c2`, inclusive."""
    for c in xrange(ord(c1), ord(c2)+1):
        yield str.capitalize(chr(c))

# Generate Part Family
def generatePartFamily():
    # Get Filename
    doc = FreeCAD.ActiveDocument
    sheet = doc.Spreadsheet
    if not doc.FileName:
        FreeCAD.Console.PrintError('Must save project first\n')
        
    docDir, docFilename = os.path.split(doc.FileName)
    filePrefix = os.path.splitext(docFilename)[0]

    columnFrom = QtGui.QInputDialog.getText(None, "Column", "Range From")
    if not columnFrom[1]:
        return
    columnTo = QtGui.QInputDialog.getText(None, "Column", "Range To")
    if not columnTo[1]:
        return
    startCell =  QtGui.QInputDialog.getInt(None, "Start Cell Row", "Input Start Cell Row:")
    if not startCell[1]:
        return
    endCell =  QtGui.QInputDialog.getInt(None, "End Cell Row", "Input End Cell Row:")
    if not endCell[1]:
        return

    fam_range = []
    for c in charRange(str(columnFrom[0]), str(columnTo[0])):
        fam_range.append(c)
    for index in range(len(fam_range)-1):
        for i in range(startCell[0],endCell[0]+1):
            cellDef = 'A'+ str(i)                        
            cellFrom = str(fam_range[index]) + str(i)
            cellTo = str(fam_range[index+1]) + str(i)
            sheet.setAlias(cellFrom, '')
            doc.recompute()
            sheet.setAlias(cellTo, sheet.getContents(cellDef))
            doc.recompute()
            sfx = str(fam_range[index+1]) + '1'
        suffix = sheet.getContents(sfx)
        
        filename = filePrefix + '_' + suffix + '.fcstd'
        filePath = os.path.join(docDir, filename)
        
        FreeCAD.Console.PrintMessage("Saving file to %s\n" % filePath)
        App.getDocument(filePrefix).saveCopy(filePath)

switch = {
    Option.SET_ALIASES: setAliases,
    Option.CLEAR_ALIASES: clearAliases,
    Option.MOVE_ALIASES: moveAliases,
    Option.GENERATE_PART_FAMILY: generatePartFamily
}                      

form = MyButtons()
form.exec_()

switch[form.retStatus]()
