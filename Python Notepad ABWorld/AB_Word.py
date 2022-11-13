from typing import Text
from PyQt5 import QtCore
from PyQt5 import QtGui
from PyQt5 import QtPrintSupport
from PyQt5.QtCore import QSize
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtPrintSupport import *
import docx2txt
import sys


class ABWord(QMainWindow):
    ''' The Main Class'''

    def __init__(self):
        super().__init__()

        #setting app logo
        self.setWindowIcon(QtGui.QIcon('AbW.png'))
        
        # setting  the geometry of window
        self.setGeometry(0, 0, 400, 300)
  
        # creating a label widget
        self.label = QLabel("Icon is set", self)
  
        # moving position
        self.label.move(100, 100)
  
        # setting up border
        self.label.setStyleSheet("border: 1px solid black;")

        # window title
        self.title = "My AB Word"
        self.setWindowTitle(self.title)

        # Editor
        self.editor = QTextEdit()
        self.setCentralWidget(self.editor)

        # menubar and toolbar
        self.create_menu_bar()
        self.create_tool_bar()
        

        # font and style
        font = QFont('Times', 15)
        self.editor.setFont(font)
        self.editor.setFontPointSize(15)

        # self.font_size_box = QSpinBox()


    
        # open size of app
        self.showMaximized()

        # path
        self.path = ''

    # def save_as_pdf(self):
    #     file_path, _ = QFileDialog.getSaveFileName(
    #         self, 'Export PDF', None, 'PDF Files (*.pdf)')
    #     printer = Qprinter(QPrinter.HighResolution)
    #     printer.setOutputFormat(QPrinter.PdfFormat)
    #     printer.setOutputFileName(file_path)
    #     self.editor.document().print_(printer)

    # menu bar items

    def create_menu_bar(self):
        menuBar = QMenuBar(self)


        # file menu
        file_menu = QMenu('File', self)
        menuBar.addMenu(file_menu)

        save_action = QAction('Save', self)
        save_action.triggered.connect(self.file_save)
        file_menu.addAction(save_action)

        open_action = QAction('Open', self)
        open_action.triggered.connect(self.file_open)
        file_menu.addAction(open_action)

        rename_action = QAction('Rename', self)
        rename_action.triggered.connect(self.file_saveas)
        file_menu.addAction(rename_action)

        pdf_action = QAction('Save as Pdf', self)
        pdf_action.triggered.connect(self.save_pdf)
        file_menu.addAction(pdf_action)

        print_action = QAction('Print', self)
        print_action.triggered.connect(self.print_widget)
        file_menu.addAction(print_action)

        # edit menu
        edit_menu = QMenu('Edit', self)
        menuBar.addMenu(edit_menu)

        # paste
        paste_action = QAction('Paste', self)
        paste_action.triggered.connect(self.editor.paste)
        edit_menu.addAction(paste_action)

        # clear
        clear_action = QAction('Clear', self)
        clear_action.triggered.connect(self.editor.clear)
        edit_menu.addAction(clear_action)

        # select all
        select_action = QAction('Select All', self)
        select_action.triggered.connect(self.editor.selectAll)
        edit_menu.addAction(select_action)

        # view menu
        view_menu = QMenu('View', self)
        menuBar.addMenu(view_menu)

        # full screen
        fullscr_action = QAction('Full Screen View', self)
        fullscr_action.triggered.connect(lambda: self.showFullScreen())
        view_menu.addAction(fullscr_action)

        # normal screen
        normscr_action = QAction('Normal View', self)
        normscr_action.triggered.connect(lambda: self.showNormal())
        view_menu.addAction(normscr_action)

        # minimize
        minscr_action = QAction('Minimize', self)
        minscr_action.triggered.connect(lambda: self.showMinimized())
        view_menu.addAction(minscr_action)

        self.setMenuBar(menuBar)

   

    def create_tool_bar(self):
        Toolbar = QToolBar("Tools", self)

        # undo
        undo_action = QAction(QIcon('undo.png'), 'undo', self)
        undo_action.triggered.connect(self.editor.undo)
        Toolbar.addAction(undo_action)

        # redo
        redo_action = QAction(QIcon('redo.png'), 'redo', self)
        redo_action.triggered.connect(self.editor.redo)
        Toolbar.addAction(redo_action)

        Toolbar.addSeparator()
        Toolbar.addSeparator()

        # cut
        cut_action = QAction(QIcon('cut.png'), 'cut', self)
        cut_action.triggered.connect(self.editor.cut)
        Toolbar.addAction(cut_action)

        # copy
        copy_action = QAction(QIcon('copy.png'), 'copy', self)
        copy_action.triggered.connect(self.editor.copy)
        Toolbar.addAction(copy_action)

        # paste
        paste_action = QAction(QIcon('paste.png'), 'paste', self)
        paste_action.triggered.connect(self.editor.paste)
        Toolbar.addAction(paste_action)

        Toolbar.addSeparator()
        Toolbar.addSeparator()


        print_action = QAction(QIcon('printing.png'),'Print', self)
        print_action.triggered.connect(self.print_widget)
        Toolbar.addAction(print_action)

        Toolbar.addSeparator()
        Toolbar.addSeparator()

        # font
        self.font_combo = QComboBox(self)
        self.font_combo.addItems(["Courier Std", "Hellentic Typewriter Regular",
                                 "Helvetica", "SanSerif", "Helvetica", "Times", "Monospace"])
        self.font_combo.activated.connect(self.set_font)
        Toolbar.addWidget(self.font_combo)

        # font size
        self.font_size = QSpinBox(self)
        self.font_size.setValue(15)
        self.font_size.valueChanged.connect(self.set_font_size)
        Toolbar.addWidget(self.font_size)

        Toolbar.addSeparator()

        # bold
        bold_action = QAction(QIcon("bold.png"), 'Bold', self)
        bold_action.triggered.connect(self.bold_text)
        Toolbar.addAction(bold_action)

        # underline
        underline_action = QAction(QIcon("underline.png"), 'Underline', self)
        underline_action.triggered.connect(self.underline_text)
        Toolbar.addAction(underline_action)

        # italic
        italic_action = QAction(QIcon("italic.png"), 'Italic', self)
        italic_action.triggered.connect(self.italic_text)
        Toolbar.addAction(italic_action)

        Toolbar.addSeparator()

        # text allignment

        right_alignment_action = QAction(
            QIcon("right-align.png"), 'Align Right', self)
        right_alignment_action.triggered.connect(
            lambda: self.editor.setAlignment(Qt.AlignRight))
        Toolbar.addAction(right_alignment_action)

        left_alignment_action = QAction(
            QIcon("left-align.png"), 'Align Left', self)
        left_alignment_action.triggered.connect(
            lambda: self.editor.setAlignment(Qt.AlignLeft))
        Toolbar.addAction(left_alignment_action)

        justification_action = QAction(
            QIcon("justification.png"), 'Center/Justify', self)
        justification_action.triggered.connect(
            lambda: self.editor.setAlignment(Qt.AlignCenter))
        Toolbar.addAction(justification_action)

        Toolbar.addSeparator()


        # zoom in
        zoom_in_action = QAction(QIcon("zoom-in.png"), 'Zoom in', self)
        zoom_in_action.triggered.connect(self.editor.zoomIn)
        Toolbar.addAction(zoom_in_action)

        # zoom out
        zoom_out_action = QAction(QIcon("zoom-out.png"), 'Zoom out', self)
        zoom_out_action.triggered.connect(self.editor.zoomOut)
        Toolbar.addAction(zoom_out_action)


        Toolbar.addSeparator()
        self.addToolBar(Toolbar)

       

    def italic_text(self):
        # if alredy italic then change to normal
        state = self.editor.fontItalic()
        self.editor.setFontItalic(not(state))

    def underline_text(self):
        # if alredy underline then change to normal
        state = self.editor.fontUnderline()
        self.editor.setFontUnderline(not(state))

    def bold_text(self):
        # if alredy bold then change to normal
        if self.editor.fontWeight() != QFont.Bold:
            self.editor.setFontWeight(QFont.Bold)
            return
        self.editor.setFontWeight(QFont.Normal)

    def set_font(self):
        font = self.font_combo.currentText()
        self.editor.setCurrentFont(QFont(font))

    def set_font_size(self):
        value = self.font_size.value()
        self.editor.setFontPointSize(value)

    def file_open(self):
        self.path, _ = QFileDialog.getOpenFileName(self, "Open file", "", "text documents (*.txt)","All files (*.*)")

        try:
            with open(self.path, 'r') as f:
                text = f.read()
            #text = docx2txt.process(self.path) for doc file
        except Exception as e:
            print(e)
        else:
            self.editor.setText(text)
            self.update_title()

    def file_save(self):
        print(self.path)
        if self.path == '':
             # If we do not have a path, we need to use Save As.
            self.file_saveas()


        text = self.editor.toPlainText()

        try:
            with open(self.path, 'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)

    def file_saveas(self):
        self.path, _ = QFileDialog.getSaveFileName(self, "Save file", "", "text documents (*.txt)","All files (*.*)")

        if self.path == '':
            return 

        text = self.editor.toPlainText()


        try:
            with open(self.path, 'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)

    def update_title(self):
        self.setWindowTitle(self.title + ' '+ self.path)

    def save_pdf(self):
        f_name, _ = QFileDialog.getSaveFileName(self, "Export PDF", None, "PDF files (*.pdf)","All files()")
        print(f_name)

        if f_name != '': # if name not empty
            printer = QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(f_name)
            self.editor.document().print_(printer)


    def print_widget(self):
        # Create printer
        printer = QPrinter()
        # Create painter
        painter = QtGui.QPainter()
        # Start painter
        painter.begin(printer)
        # Grab a widget you want to print
        screen = self.editor.grab()
        # Draw grabbed pixmap
        painter.drawPixmap(10, 10, screen)
        # End painting
        painter.end()




app = QApplication(sys.argv)
window = ABWord()
window.show()
sys.exit(app.exec_())


#AbWorld2