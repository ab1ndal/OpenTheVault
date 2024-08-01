import sys
from PyQt5.QtWidgets import QButtonGroup, QCheckBox, QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton, QFormLayout, QComboBox, QDoubleSpinBox, QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox, QRadioButton
from PyQt5.QtGui import QFont
from createCut import CreateCut


class SectionCutForm(QWidget):
    def __init__(self):
        super().__init__()
        self.sectionCut = CreateCut()

        self.initUI()

    def initUI(self):
        print("Initializing UI")
        self.createWidgets()
        print("Widgets Created")
        self.setupLayout()
        self.setupConnections()

    def createWidgets(self):
        # Font setup
        font = QFont('Arial', 12)  # Example: Arial font, size 12 points

        # Form elements (combined dictionary)
        self.widgets = {
            'cutName': {'widget': QLineEdit(), 'label': QLabel('Section Cut Prefix:')},
            'cutDirection': {'widget': QComboBox(), 'label': QLabel('Normal Direction:')},
            'groupName': {'widget': QLineEdit(), 'label': QLabel('Section Cut Group:')},
            'startCoord': {'widget': QDoubleSpinBox(), 'label': QLabel('Normal Start Coordinate:')},
            'endCoord': {'widget': QDoubleSpinBox(), 'label': QLabel('Normal End Coordinate:')},
            'cutStep': {'widget': QDoubleSpinBox(), 'label': QLabel('Cut Spacing:')},
            'pointGroup_2Pt':{'widget': QRadioButton('Diagonal Points'), 'label': QLabel('Rectangle:')},
            'addSpecialCoord':{'widget':QLineEdit(), 'label':QLabel('Add cuts at Special Coordinates:')},
            'rmSpecialCoord':{'widget':QLineEdit(), 'label':QLabel('Remove cuts at Special Coordinates:')},
            'pointGroup_4Pt':{'widget': QRadioButton('4 Points'), 'label': QLabel('')},
            'x1': {'widget': QDoubleSpinBox(), 'label': QLabel('X1:')},
            'y1': {'widget': QDoubleSpinBox(), 'label': QLabel('Y1:')},
            'x2': {'widget': QDoubleSpinBox(), 'label': QLabel('X2:')},
            'y2': {'widget': QDoubleSpinBox(), 'label': QLabel('Y2:')},
            'x3': {'widget': QDoubleSpinBox(), 'label': QLabel('X3:')},
            'y3': {'widget': QDoubleSpinBox(), 'label': QLabel('Y3:')},
            'x4': {'widget': QDoubleSpinBox(), 'label': QLabel('X4:')},
            'y4': {'widget': QDoubleSpinBox(), 'label': QLabel('Y4:')},
            'advAxisExists': {'widget': QCheckBox(), 'label': QLabel('Advanced Axis:')},
            'localPlane': {'widget': QComboBox(), 'label': QLabel('Local Plane:')},
            'vec1': {'widget': QLineEdit(), 'label': QLabel('Plane Joint 1:')},
            'vec2': {'widget': QLineEdit(), 'label': QLabel('Plane Joint 2:')},
            'unit': {'widget': QLineEdit(), 'label': QLabel('Units:')},
            'folderLocation': {'widget': QLineEdit(), 'label': QLabel('Save Excel in Folder:')}
        }
            
        self.buttons = {
            'browseButton': {'widget': QPushButton('Browse'), 'label': QLabel('')},  # Empty label for button
            'addCutButton': {'widget': QPushButton('Add Cut'), 'label': QLabel('')},  # Empty label for button
            'exportExcelButton': {'widget': QPushButton('Export as Excel'), 'label': QLabel('')},  # Empty label for button
            'message': {'widget': QLabel(), 'label': QLabel('')}  # Empty label for message label
        }

        # Set fonts for all widgets
        self.setFonts(font, [widget['widget'] for widget in self.widgets.values() if isinstance(widget['widget'], (QLineEdit, QComboBox, QDoubleSpinBox, QLabel, QCheckBox))])
        self.setFonts(font, [widget['label'] for widget in self.widgets.values() if isinstance(widget['label'], (QLabel))])
        self.setFonts(font, [button['widget'] for button in self.buttons.values() if isinstance(button['widget'], (QPushButton, QLabel))])
        # Set font for radio buttons
        self.widgets['pointGroup_2Pt']['widget'].setFont(font)
        self.widgets['pointGroup_4Pt']['widget'].setFont(font)

        # Populate combo box
        self.widgets['cutDirection']['widget'].addItems(['X', 'Y', 'Z'])
        self.widgets['localPlane']['widget'].addItems([f'{i}-{j}' for i in range(1,4) for j in range(1,4) if i!=j])
        # set default values
        self.widgets['cutDirection']['widget'].setCurrentIndex(2)
        self.widgets['localPlane']['widget'].setCurrentIndex(4)

        # Define values for Radio Button
        self.pointGroup = QButtonGroup()
        self.pointGroup.addButton(self.widgets['pointGroup_2Pt']['widget'])
        self.pointGroup.addButton(self.widgets['pointGroup_4Pt']['widget'])
        self.widgets['pointGroup_2Pt']['widget'].toggled.connect(self.onRadioButtonToggled)
        
        self.advAxisChecked()
        

        # Allow negative values for QDoubleSpinBox
        for key in ['startCoord', 'endCoord', 'cutStep', 'x1', 'y1', 'x2', 'y2', 'x3', 'y3', 'x4', 'y4']:
            self.widgets[key]['widget'].setMinimum(-999999999)
            self.widgets[key]['widget'].setMaximum(999999999)
            self.widgets[key]['widget'].setDecimals(3)

    def setFonts(self, font, widgets):
        for widget in widgets:
            widget.setFont(font)

    def setupLayout(self):
        layout = QVBoxLayout()
        form_layout = QFormLayout()
        

        # Add form elements
        for key, value in self.widgets.items():
            if value['label']:
                if key == 'folderLocation':
                    folder_layout = QHBoxLayout()
                    folder_layout.addWidget(value['widget'])
                    folder_layout.addWidget(self.buttons['browseButton']['widget'])
                    form_layout.addRow(value['label'], folder_layout)
                
                elif key == 'pointGroup_2Pt':
                    radio_layout = QHBoxLayout()
                    radio_layout.addWidget(value['widget'])
                    radio_layout.addWidget(self.widgets['pointGroup_4Pt']['widget'])
                    form_layout.addRow(value['label'], radio_layout)
                
                elif key == 'pointGroup_4Pt':
                    continue
                
                else:
                    form_layout.addRow(value['label'], value['widget'])

        layout.addLayout(form_layout)

        # Add buttons and message label
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.buttons['addCutButton']['widget'])
        button_layout.addWidget(self.buttons['exportExcelButton']['widget'])
        layout.addLayout(button_layout)
        layout.addWidget(self.buttons['message']['widget'])

        self.setLayout(layout)

    def setupConnections(self):
        self.widgets['cutDirection']['widget'].currentIndexChanged.connect(self.updateLabels)
        self.buttons['browseButton']['widget'].clicked.connect(self.browseFolder)
        self.buttons['addCutButton']['widget'].clicked.connect(self.addCut)
        self.buttons['exportExcelButton']['widget'].clicked.connect(self.exportExcel)
        #If Advanced Axis is checked, enable the Local Plane, vec1 and vec2 fields, else disable them
        self.widgets['advAxisExists']['widget'].stateChanged.connect(self.advAxisChecked)

    def advAxisChecked(self):
        if self.widgets['advAxisExists']['widget'].isChecked():
            self.widgets['localPlane']['widget'].setEnabled(True)
            self.widgets['vec1']['widget'].setEnabled(True)
            self.widgets['vec2']['widget'].setEnabled(True)
        else:
            self.widgets['localPlane']['widget'].setEnabled(False)
            self.widgets['vec1']['widget'].setEnabled(False)
            self.widgets['vec2']['widget'].setEnabled(False)

    def updateLabels(self):
        direction = self.widgets['cutDirection']['widget'].currentText()
        labels_mapping = {
            'X': ['Y1:', 'Z1:', 'Y2:', 'Z2:', 'Y3:', 'Z3:', 'Y4:', 'Z4:'],
            'Y': ['X1:', 'Z1:', 'X2:', 'Z2:', 'X3:', 'Z3:', 'X4:', 'Z4:'],
            'Z': ['X1:', 'Y1:', 'X2:', 'Y2:', 'X3:', 'Y3:', 'X4:', 'Y4:']
        }

        for index, label in enumerate(['x1', 'y1', 'x2', 'y2', 'x3', 'y3', 'x4', 'y4']):
            self.widgets[label]['label'].setText(labels_mapping[direction][index])

    def readCoordinates(self,text):
        if text == '':
            return []
        coordList = text.split(',')
        return [round(float(x),3) for x in coordList]
    
    def onRadioButtonToggled(self):
        if self.widgets['pointGroup_2Pt']['widget'].isChecked():
            self.widgets['x3']['widget'].setEnabled(False)
            self.widgets['y3']['widget'].setEnabled(False)
            self.widgets['x4']['widget'].setEnabled(False)
            self.widgets['y4']['widget'].setEnabled(False)
        else:
            self.widgets['x3']['widget'].setEnabled(True)
            self.widgets['y3']['widget'].setEnabled(True)
            self.widgets['x4']['widget'].setEnabled(True)
            self.widgets['y4']['widget'].setEnabled(True)
            
    def browseFolder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Folder')
        if folder:
            self.widgets['folderLocation']['widget'].setText(folder)
    
    def validateForm(self):
        for key, value in self.widgets.items():
            if isinstance(value['widget'], (QLineEdit)):
                if value['widget'].text() == '' and value['widget'].isEnabled():
                    return False
            elif isinstance(value['widget'], QComboBox):
                if value['widget'].currentText() == '':
                    return False
        return True
    
    def assignValues(self):
        self.sectionCut.inputCutName(cutName=self.widgets['cutName']['widget'].text())
        if self.widgets['pointGroup_2Pt']['widget'].isChecked():
            self.sectionCut.inputIs4Pt(is4Pt = False)
            maxX = max(self.widgets['x1']['widget'].value(), self.widgets['x2']['widget'].value()) 
            minX = min(self.widgets['x1']['widget'].value(), self.widgets['x2']['widget'].value())
            maxY = max(self.widgets['y1']['widget'].value(), self.widgets['y2']['widget'].value())
            minY = min(self.widgets['y1']['widget'].value(), self.widgets['y2']['widget'].value())
            self.sectionCut.inputDiagCoord(diagCoord=[maxX, maxY, minX, minY]) 
        self.sectionCut.inputCutDirection(cutDirection=self.widgets['cutDirection']['widget'].currentText())
        self.sectionCut.inputCutStep(cutStep=self.widgets['cutStep']['widget'].value())
        self.sectionCut.inputStartCoord(startCoord=self.widgets['startCoord']['widget'].value())
        self.sectionCut.inputEndCoord(endCoord=self.widgets['endCoord']['widget'].value())
        self.sectionCut.inputGroupName(groupName=self.widgets['groupName']['widget'].text())
        self.sectionCut.inputUnit(unit=self.widgets['unit']['widget'].text())
        self.sectionCut.inputAdvAxis(advAxisExists=self.widgets['advAxisExists']['widget'].isChecked())
        self.sectionCut.inputSpecialCoord(addSpecialCoord=self.readCoordinates(self.widgets['addSpecialCoord']['widget'].text()),
                                          rmSpecialCoord=self.readCoordinates(self.widgets['rmSpecialCoord']['widget'].text()))
        if self.widgets['advAxisExists']['widget'].isChecked():
            self.sectionCut.inputVec1(vec1=self.widgets['vec1']['widget'].text())
            self.sectionCut.inputVec2(vec2=self.widgets['vec2']['widget'].text())
            self.sectionCut.inputLocalPlane(localPlane=self.widgets['localPlane']['widget'].currentText())

        if self.widgets['pointGroup_4Pt']['widget'].isChecked():
            self.sectionCut.inputIs4Pt(is4Pt = True)
            self.sectionCut.input4PtCoord(x1=self.widgets['x1']['widget'].value(),
                                     y1=self.widgets['y1']['widget'].value(),
                                     x2=self.widgets['x2']['widget'].value(),
                                     y2=self.widgets['y2']['widget'].value(),
                                     x3=self.widgets['x3']['widget'].value(), 
                                     y3=self.widgets['y3']['widget'].value(), 
                                     x4=self.widgets['x4']['widget'].value(), 
                                     y4=self.widgets['y4']['widget'].value())
        return

    def addCut(self):
        # Check if all fields are filled
        cutName = self.widgets['cutName']['widget'].text()
        if not self.validateForm():
            self.showMessage('Please fill all the fields.')
            return
        
        self.assignValues()
        self.sectionCut.defineCut()

        # Process form data or perform additional validation
        self.showMessage(f"Cut '{cutName}' was added.")

    def exportExcel(self):
        folder = self.widgets['folderLocation']['widget'].text()
        if not folder:
            self.showMessage('Please select a folder.')
            return
        self.sectionCut.printExcel(fileLoc=folder)

        # Perform export operations using folder path
        self.showMessage('Data exported as Excel.')

    def showMessage(self, message):
        self.buttons['message']['widget'].setText(message)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Set Fusion style
    window = QMainWindow()
    main_widget = SectionCutForm()
    window.setCentralWidget(main_widget)
    window.setWindowTitle('Section Cut Form for SAP2000')
    window.show()
    sys.exit(app.exec_())
    print("App Started")
