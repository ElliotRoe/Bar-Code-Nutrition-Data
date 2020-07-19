from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import QtGui
import qdarkgraystyle
from datetime import date
import requests
import threading
import os
import bar_search
import traceback, sys

# Defines all the signals we'll need for the the various processes we are running in Qthreads
class WorkerSignals(QObject):
    finished = pyqtSignal()
    result = pyqtSignal(object)
    error = pyqtSignal(tuple)
    progress = pyqtSignal(int)

class Worker(QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()

        # Store constructor arguments (re-used for processing)
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

        # Add the callback to our kwargs
        self.kwargs['progress_callback'] = self.signals.progress

    @pyqtSlot()
    def run(self):
        # Retrieve args/kwargs here; and fire processing using them
        try:
            result = self.fn(*self.args, **self.kwargs)
        except:
            traceback.print_exc()
            exctype, value = sys.exc_info()[:2]
            self.signals.error.emit((exctype, value, traceback.format_exc()))
        else:
            self.signals.result.emit(result)  # Return the result of the processing
        finally:
            self.signals.finished.emit()  # Done


class Main_Window(QMainWindow):
    def __init__(self, *args, **kwargs):
        super(Main_Window, self).__init__(*args, **kwargs)

        # Download thread
        self.download = None
        self.window_exited = False
        self.downloading = False;

        # Creates threadpool to handle threads
        self.threadpool = QThreadPool()

        # Scraping thread
        self.scrape = threading.Thread(target=self.scrape_data)

        self.fileName = ''

        self.style = qdarkgraystyle.load_stylesheet()

        self.setStyleSheet(self.style)
        self.setFixedSize(350, 220)
        self.setWindowTitle("Bar Code Analyzer")
        self.setWindowIcon(QtGui.QIcon("graphics/blue-document-excel-table.png"))

        self.main_layout = QVBoxLayout()

        self.top_layout = QVBoxLayout()

        self.choose_layout = QHBoxLayout()

        self.update_layout = QHBoxLayout()

        self.output_file_layout = QHBoxLayout()

        self.control_layout = QHBoxLayout()

        # Choose button horizontal layout
        self.choose_button = QPushButton('Choose File')
        self.choose_button.setFixedWidth(75)
        self.choose_button.clicked.connect(self.choose_file)
        self.choose_button.setToolTip('Choose input .xlsx file for barcode analyzer')

        self.hchoose_button = QPushButton('?')
        self.hchoose_button.clicked.connect(
            lambda: self.info_popup('About the Input File', 'Please select the excel file '
                                                            'containing the barcodes you '
                                                            'wish to input \n\nThe '
                                                            'requirements for this file '
                                                            'are as follows:\n1. The '
                                                            'barcodes must be stored in '
                                                            'the first column.\n2. Column '
                                                            'A must have a header of '
                                                            '"Barcodes".\n3. The barcode '
                                                            'data must be stored as a '
                                                            'String.'))
        self.hchoose_button.setFixedWidth(25)
        self.hchoose_button.setToolTip('Help')

        self.choose_label = QLabel()

        self.choose_layout.addWidget(self.choose_button)
        self.choose_layout.addWidget(self.hchoose_button)
        self.choose_layout.addWidget(self.choose_label)

        # Update button horizontal layout

        self.update_button = QPushButton('Update')
        self.update_button.clicked.connect(self.update_database)
        self.update_button.setToolTip('Updates database to current version')

        self.hupdate_button = QPushButton('?')
        self.hupdate_button.setFixedWidth(25)
        self.hupdate_button.clicked.connect(lambda: self.info_popup('About Updating', 'This functions updates the '
                                                                                      '.csv file used for pulling the '
                                                                                      'data for each for of the '
                                                                                      'barcodes. It takes about 2-3 '
                                                                                      'minutes to update and is '
                                                                                      'generally a good idea if you '
                                                                                      'think some data is missing in '
                                                                                      'the output. Please do not '
                                                                                      'close the program while it is '
                                                                                      'updating'))
        self.hupdate_button.setToolTip('Help')

        date_txt = open('data/update_date.txt', 'r')
        date = date_txt.readline()
        self.update_label = QLabel('Last Updated: ' + date[:len(date) - 1])
        self.update_label.setAlignment(Qt.AlignCenter)

        self.update_layout.addWidget(self.update_button)
        self.update_layout.addWidget(self.hupdate_button)
        self.update_layout.addWidget(self.update_label)

        # Creates Scrape Data button

        self.scrape_button = QPushButton('Scrape Data')
        self.scrape_button.setEnabled(False)
        self.scrape_button.clicked.connect(self.scrape_wrapper)

        # Output file path

        self.output_label = QLabel('Output Path:')
        self.output_entry = QLineEdit()
        self.output_entry.setText(os.path.expanduser('~\\Documents\\output_workbook.xlsx'))

        self.output_file_layout.addWidget(self.output_label)
        self.output_file_layout.addWidget(self.output_entry)

        # Assembles top layout

        self.top_layout.addLayout(self.choose_layout)
        self.top_layout.addLayout(self.update_layout)
        self.top_layout.addSpacerItem(QSpacerItem(10,5))
        self.top_layout.addLayout(self.output_file_layout)
        self.top_layout.addWidget(self.scrape_button)

        # Assembles main layout

        self.main_layout.addLayout(self.top_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setValue(0)
        self.main_layout.addWidget(self.progress_bar)

        # Sets the layout to main window

        self.widget = QWidget()
        self.widget.setLayout(self.main_layout)

        self.setCentralWidget(self.widget)

    def closeEvent(self, event):
        self.window_exited = True
        self.downloading = False

        print('Quitting')
        self.download.quit()
        os.remove('data/temp.csv')
        if self.scrape.is_alive():
            self.scrape.join()

    def choose_file(self):
        self.fileName = QFileDialog.getOpenFileName(self, 'Select Input Excel File', os.path.expanduser('~/Documents'), 'Excel Files (*.xlsx)')
        temp = str(self.fileName[0])
        try:
            temp.index(':')
        except ValueError:
            return
        else:
            self.choose_label.setText(temp)
            self.scrape_button.setEnabled(True)

    def throw_error(self, title, text):
        error = QMessageBox()
        error.setWindowTitle(title)
        error.setText(text)
        error.setWindowIcon(QtGui.QIcon("graphics/compile-error.png"))
        error.setIcon(QMessageBox.Critical)

        error.exec()

    def info_popup(self, title, text):
        popup = QMessageBox()
        popup.setWindowTitle(title)
        popup.setText(text)
        popup.setWindowIcon(QtGui.QIcon("graphics/information.png"))
        popup.setIcon(QMessageBox.Information)

        popup.exec_()

    def update_database(self):
        today = date.today()
        self.update_date = today.strftime("%m/%d/%y")
        self.update_label.setText('Last Updated: ' + self.update_date)

        date_txt = open("data/update_date.txt", "w")

        print(self.update_date, file=date_txt)
        date_txt.close()

        self.update_button.setEnabled(False)
        self.scrape_button.setEnabled(False)

        self.download = Worker(self.download_csv)

        self.downloading = True;

        self.download.signals.finished.connect(self.download_cleanup)
        self.download.signals.progress.connect(self.progress_bar.setValue)

        self.threadpool.start(self.download)

    def download_csv(self, progress_callback):
        print('Beginning file download with requests')

        url = "https://static.openfoodfacts.org/data/en.openfoodfacts.org.products.csv"

        # Downloading file
        with open('/Users/Elliot/PycharmProjects/Bar-Code-Nutrition-Data/data/temp.csv',
                  'wb') as f:
            response = requests.get(url, stream=True)
            total_length = 2917101258
            dl = 0
            total_length = int(total_length)
            for data in response.iter_content(chunk_size=8192):
                dl += len(data)
                f.write(data)

                # Calculating approx progress
                progress = int((float(dl) / float(total_length)) * 100)

                # Emitting signal to progress update function in order to safely update the value
                progress_callback.emit(progress)

                # Abort flag for thread in case download gets interrupted
                #if not self.downloading:
                    #return
            f.close()

    def download_cleanup(self):
        print('Clean starting')
        os.remove('data/en.openfoodfacts.org.products.csv')
        os.rename('data/temp.csv', 'data/en.openfoodfacts.org.products.csv')
        self.reset_window()

    def scrape_wrapper(self):
        if not self.scrape.is_alive():
            self.update_button.setEnabled(False)
            self.choose_button.setEnabled(False)
            self.scrape.start()
            self.scrape_button.setText('Stop')
        else:
            self.window_exited = True
            self.scrape.join()
            self.reset_window()



    def scrape_data(self):
        bc = bar_search.getBarcodes(self.fileName[0])
        scraped_data = bar_search.read_csv(bc, window=self)
        if scraped_data is None:
            return
        bar_search.write_xl(scraped_data, path=self.output_entry.text())
        return

    def reset_window(self):
        self.scrape_button.setEnabled(True)
        self.scrape_button.setText('Scrape Data')
        self.choose_button.setEnabled(True)
        self.progress_bar.setValue(0)


app = QApplication([])

window = Main_Window()
window.show()

app.exec_()
