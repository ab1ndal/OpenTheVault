import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton, QFormLayout, QComboBox, QDoubleSpinBox, QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox
from PyQt5.QtGui import QFont
from createCut import CreateCut


class SectionCutForm(QWidget):
    def __init__(self):
        super().__init__()
        self.sectionCut = CreateCut()

        self.initUI()

    def initUI(self):
        self.createWidgets()
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
            'x1': {'widget': QDoubleSpinBox(), 'label': QLabel('X1:')},
            'y1': {'widget': QDoubleSpinBox(), 'label': QLabel('Y1:')},
            'x2': {'widget': QDoubleSpinBox(), 'label': QLabel('X2:')},
            'y2': {'widget': QDoubleSpinBox(), 'label': QLabel('Y2:')},
            'vec1': {'widget': QLineEdit(), 'label': QLabel('Vector 1:')},
            'vec2': {'widget': QLineEdit(), 'label': QLabel('Vector 2:')},
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
        self.setFonts(font, [widget['widget'] for widget in self.widgets.values() if isinstance(widget['widget'], (QLineEdit, QComboBox, QDoubleSpinBox, QLabel))])
        self.setFonts(font, [widget['label'] for widget in self.widgets.values() if isinstance(widget['label'], (QLabel))])
        self.setFonts(font, [button['widget'] for button in self.buttons.values() if isinstance(button['widget'], (QPushButton, QLabel))])

        # Populate combo box
        self.widgets['cutDirection']['widget'].addItems(['X', 'Y', 'Z'])
        # set default values
        self.widgets['cutDirection']['widget'].setCurrentIndex(2)

        # Allow negative values for QDoubleSpinBox
        for key in ['startCoord', 'endCoord', 'cutStep', 'x1', 'y1', 'x2', 'y2']:
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

    def updateLabels(self):
        direction = self.widgets['cutDirection']['widget'].currentText()
        labels_mapping = {
            'X': ['Y1:', 'Z1:', 'Y2:', 'Z2:'],
            'Y': ['X1:', 'Z1:', 'X2:', 'Z2:'],
            'Z': ['X1:', 'Y1:', 'X2:', 'Y2:']
        }

        for index, label in enumerate(['x1', 'y1', 'x2', 'y2']):
            self.widgets[label]['label'].setText(labels_mapping[direction][index])
            
    def browseFolder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Folder')
        if folder:
            self.widgets['folderLocation']['widget'].setText(folder)
    
    def validateForm(self):
        for key, value in self.widgets.items():
            if isinstance(value['widget'], (QLineEdit)):
                if value['widget'].text() == '':
                    return False
            elif isinstance(value['widget'], QComboBox):
                if value['widget'].currentText() == '':
                    return False
        return True
    
    def assignValues(self):
        self.sectionCut.inputCutName(cutName=self.widgets['cutName']['widget'].text())
        self.sectionCut.inputDiagCoord(diagCoord=[
            self.widgets['x1']['widget'].value(),
            self.widgets['y1']['widget'].value(),
            self.widgets['x2']['widget'].value(),
            self.widgets['y2']['widget'].value()
        ]) 
        self.sectionCut.inputCutDirection(cutDirection=self.widgets['cutDirection']['widget'].currentText())
        self.sectionCut.inputCutStep(cutStep=self.widgets['cutStep']['widget'].value())
        self.sectionCut.inputStartCoord(startCoord=self.widgets['startCoord']['widget'].value())
        self.sectionCut.inputEndCoord(endCoord=self.widgets['endCoord']['widget'].value())
        self.sectionCut.inputGroupName(groupName=self.widgets['groupName']['widget'].text())
        self.sectionCut.inputUnit(unit=self.widgets['unit']['widget'].text())
        self.sectionCut.inputVec1(vec1=self.widgets['vec1']['widget'].text())
        self.sectionCut.inputVec2(vec2=self.widgets['vec2']['widget'].text())
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
