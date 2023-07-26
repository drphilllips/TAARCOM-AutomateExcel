import os
import sys

from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QDialog, QApplication, QFileDialog
from PyQt5.uic import loadUi

import Clean
import Compile
import EnumTypes
import Split

VERSION = "Master v1.0.0"


class Stream(QtCore.QObject):
    """Redirects console output to text widget"""
    newText = QtCore.pyqtSignal(str)

    def write(self, text):
        self.newText.emit(str(text))

    # Pass the flush so we don't get an attribute error.
    def flush(self):
        pass


class MainWindow(QDialog):
    """Generates the main window for our program"""

    def __init__(self):
        super(MainWindow, self).__init__()

        # External UI design w/ QTDesigner ;)
        loadUi("AutomateExcel.ui", self)

        # Initialize the threadpool for handling worker jobs
        self.threadpool = QtCore.QThreadPool()

        # State variables
        self.filepaths = []

        # Connect GUI buttons to methods
        self.btnSelectFiles.clicked.connect(self.selectFiles)
        self.btnDeselectFiles.clicked.connect(self.deselectFiles)
        self.btnClearConsole.clicked.connect(self.clearConsole)
        self.btnClean.clicked.connect(self.cleanClicked)
        self.btnSplit.clicked.connect(self.splitClicked)
        self.btnCompile.clicked.connect(self.compileClicked)
        # Disable Excel operations until file selected
        self.btnClean.setEnabled(False)
        self.btnSplit.setEnabled(False)
        self.btnCompile.setEnabled(False)

        # Custom output stream
        sys.stdout = Stream(newText=self.writeToConsole)

        # Show welcome message
        self.clearConsole()
        print("> Make sure to pull the latest version from GitHub!")

    # ------------------------
    #  One Excel Op at a Time
    # ------------------------

    def cleanClicked(self):
        """Send the Clean execution to a worker thread."""
        self.lockButtons()
        worker = Worker(self.clean)
        if self.threadpool.activeThreadCount() == 0:
            self.threadpool.start(worker)

    def splitClicked(self):
        """Send the Split execution to a worker thread."""
        self.lockButtons()
        worker = Worker(self.split)
        if self.threadpool.activeThreadCount() == 0:
            self.threadpool.start(worker)

    def compileClicked(self):
        """Send the Compile execution to a worker thread."""
        self.lockButtons()
        worker = Worker(self.compile)
        if self.threadpool.activeThreadCount() == 0:
            self.threadpool.start(worker)

    # -------------------------
    #  Execute Operation Files
    # -------------------------

    def clean(self):
        """Runs function for Clean"""

        # Only run on one file at a time
        # if len(self.filepaths) > 1:
        #     first_filepath = self.filepaths[0]
        #     self.deselectFiles()
        #     self.filepaths = [first_filepath]
        #     print("..Only cleaning the first selected file.")

        # Check if we have the necessary lookup files
        look_dir = "W:/Lookup/"
        rcl_exists = os.path.exists(look_dir + "RootColumnLibrary.xlsx")  # Root Column Library
        cpm_exists = os.path.exists(look_dir + "rootCustomerMappings.xlsx")  # Customer-ProperName Map
        mal_exists = os.path.exists(look_dir + "Master Account List.xlsx")  # Master Account List
        mtl_exists = os.path.exists(look_dir + "CAZipCode.xlsx")  # Master Territory List

        # Only run if all lookup files can be found
        if rcl_exists and cpm_exists and mal_exists and mtl_exists:
            # Run the Clean.py file.
            try:
                # Get company that produced this insight file
                company_txt = self.comboCompany.currentText()
                company = EnumTypes.Company.NA
                if company_txt == "DGK":
                    company = EnumTypes.Company.DGK
                elif company_txt == "MOU":
                    company = EnumTypes.Company.MOU
                elif company_txt == "ABR":
                    company = EnumTypes.Company.ABR
                # Clean all selected files
                for filepath in self.filepaths:
                    # Strip root off path to get file name
                    filename = os.path.basename(filepath)
                    # Make sure the file is not already standardized
                    if "Standardized" in filename:
                        print(".." + filename + " has already been cleaned.")
                        self.unlockButtons()
                    else:
                        Clean.main(filepath, company)
            except Exception as error:
                print("..Unexpected Python error:\n" +
                      "?" + str(error) + "\n" +
                      "..Please contact your local coder.")
            # Clear file.
            self.unlockButtons()
            self.deselectFiles()
        elif not self.filepaths:
            print("..No Insight files selected!\n"
                  "..Use the Select File button to select files.")
        elif not rcl_exists:
            print("..File RootColumnLibrary.xlsx not found!\n"
                  "..Please check file location and try again.")
        elif not cpm_exists:
            print("..File rootCustomerMappings.xlsx not found!\n"
                  "..Please check file location and try again.")
        elif not mal_exists:
            print("..File Master Account List.xlsx not found!\n"
                  "..Please check file location and try again.")
        elif not mtl_exists:
            print("..File CAZipCode.xlsx not found!\n"
                  "..Please check file location and try again.")

    def split(self):
        """Runs function for split"""

        # Only run one file at a time
        if len(self.filepaths) > 1:
            first_filepath = self.filepaths[0]
            self.deselectFiles()
            self.filepaths = [first_filepath]
            print("..Only cleaning the first selected file.")

        # Make sure the file is standardized
        filename = os.path.basename(self.filepaths[0])
        if "Standardized" not in filename:
            print("..Split can only be run on standardized files.\n"
                  "..Make sure the filename contains \"Standardized\".")
            self.unlockButtons()
            return

        # Run the Split.py file
        try:
            Split.main(self.filepaths[0])
        except Exception as error:
            print("..Unexpected Python error:\n" +
                  "?" + str(error) + "\n" +
                  "..Please contact your local coder.")
        # Clear file.
        self.unlockButtons()
        self.deselectFiles()

    def compile(self):
        """Runs function for compile"""

        # Make sure there are multiple files
        if len(self.filepaths) < 2:
            print("..Compiling requires multiple files.")
            self.unlockButtons()
            return

        # Run the Compile.py file
        try:
            Compile.main(self.filepaths)
        except Exception as error:
            print("..Unexpected Python error:\n" +
                  "?" + str(error) + "\n" +
                  "..Please contact your local coder.")
        # Clear files
        self.unlockButtons()
        self.deselectFiles()

    # -----------------------
    #  GUI Utility Functions
    # -----------------------

    def lockButtons(self):
        """Disable user interaction"""

        self.btnSelectFiles.setEnabled(False)
        self.btnDeselectFiles.setEnabled(False)
        self.btnClearConsole.setEnabled(False)
        self.btnClean.setEnabled(False)
        self.btnSplit.setEnabled(False)
        self.btnCompile.setEnabled(False)
        self.comboCompany.setEnabled(False)

    def unlockButtons(self):
        """Enable user interaction"""

        self.btnSelectFiles.setEnabled(True)
        self.btnDeselectFiles.setEnabled(True)
        self.btnClearConsole.setEnabled(True)
        self.btnClean.setEnabled(True)
        self.btnSplit.setEnabled(True)
        self.btnCompile.setEnabled(True)
        self.comboCompany.setEnabled(True)

    def writeToConsole(self, text):
        """Write console output to text widget."""

        cursor = self.txtConsole.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.txtConsole.setTextCursor(cursor)
        self.txtConsole.ensureCursorVisible()

    def clearConsole(self):
        """Clear console print statements"""

        self.txtConsole.clear()
        print("> Welcome to the TAARCOM, Inc. Automate Excel Program!")

    def selectFiles(self):
        """Select file for Excel operations"""

        # Let user know the old selection is cleared
        if self.filepaths:
            self.filepaths = []
            print("..Selecting new file, old selection cleared..")

        # Clear "<No Files Selected>"
        self.lblSelectedFiles.setText("")

        # Grab Excel file for operations
        self.filepaths, _ = QFileDialog.getOpenFileNames(
            self, filter="Excel files (*.xls *.xlsx *.xlsm)")

        # Print out the selected filenames
        for filename in [os.path.basename(filepath) for filepath in self.filepaths]:
            print("> File selected: " + filename)
            # Shorten filename if too long for selected files label
            if len(filename) > 43:
                filename = filename[:43] + "..."
            # Update current file label
            self.lblSelectedFiles.setText(self.lblSelectedFiles.text() + "> " + filename + "\n")

        # Enable Excel operations now that file is selected
        self.btnClean.setEnabled(True)
        self.btnSplit.setEnabled(True)
        self.btnCompile.setEnabled(True)

    def deselectFiles(self):
        """Deselect files and adjust GUI accordingly"""

        if self.filepaths:
            self.filepaths = []
            self.lblSelectedFiles.setText("<No Files Selected>")
            print("> File selection cleared.")

        # Disable Excel operations now that file is deselected
        self.btnClean.setEnabled(False)
        self.btnSplit.setEnabled(False)
        self.btnCompile.setEnabled(False)


class Worker(QtCore.QRunnable):
    """Inherits from QRunnable to handle worker thread.

    param args -- Arguments to pass to the callback function.
    param kwargs -- Keywords to pass to the callback function.
    """
    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()
        # Store constructor arguments (re-used for processing)
        self.fn = fn
        self.args = args
        self.kwargs = kwargs

    @pyqtSlot()
    def run(self):
        """Initialize the runner function with passed args, kwargs."""
        self.fn(*self.args, **self.kwargs)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # widget container for QT Designer UI
    widget = QtWidgets.QStackedWidget()
    widget.setWindowTitle("Automate Excel (" + VERSION + ")")
    main_window = MainWindow()
    widget.addWidget(main_window)
    widget.setFixedWidth(900)
    widget.setFixedHeight(600)
    widget.show()

try:
    sys.exit(app.exec_())
except:
    print("..Exiting")
