from openpyxl import load_workbook, Workbook, worksheet
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.formatting.rule import ColorScaleRule
from PyQt5.QtWidgets import (QApplication, QFileDialog, QWidget, QPushButton, QDesktopWidget, QGridLayout)
import string
import pandas as pd
import calendar
import datetime
import numpy as np
import shelve
import sys
import os
import re

files = ['Tender',
         'Employee Sale',
         'Tax Free Sale',
         'Purchased GC',
         'Redeemed GC',
         'CM Report',
         'FOLDER CONTAINING ALL FILES...']

coordinates = [(x, y) for x in range(len(files)) for y in range(1)]
FileLocations = {'File Name': [], 'Location': []}


class MainWindow(QWidget):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.title = 'DSR File Selector'
        self.left = 10
        self.top = 10
        self.width = 320
        self.height = 200
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        # Moving Window to the center of the screen
        qtRectangle = self.frameGeometry()
        centerPoint = QDesktopWidget().availableGeometry().center()
        qtRectangle.moveCenter(centerPoint)
        self.move(qtRectangle.topLeft())

        grid = QGridLayout()
        self.setLayout(grid)
        for coordinate, file in zip(coordinates, files):
            if file == '':
                continue
            if file == 'FOLDER CONTAINING ALL FILES...':
                fileDirectoryButton = QPushButton(file)
                grid.addWidget(fileDirectoryButton, *coordinate)
                fileDirectoryButton.clicked.connect(self.get_directory)

            else:
                fileSearchButton = QPushButton(file)
                grid.addWidget(fileSearchButton, *coordinate)
                fileSearchButton.clicked.connect(self.search_file)

    def search_file(self):
        options = QFileDialog.Options()
        find_file, _ = QFileDialog.getOpenFileName(self, 'DSR File 1', '',
                                                   'Excel Files (*.xlsx *xls)',
                                                   options=options)
        FileLocations['File Name'].append(self.sender().text())
        FileLocations['Location'].append(find_file)

    def get_directory(self):
        dialog = QFileDialog()
        folder_path = dialog.getExistingDirectory(self, None, 'Select Folder with ALL 6 Files')
        FileLocations['File Name'].append('Directory')
        FileLocations['Location'].append(folder_path)
        dialog.setEnabled(False)


def file_selector():
    if __name__ == '__main__':
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        app.exec_()


def missing_tables(df):
    try:
        df
        return df
    except NameError:
        df = pd.DataFrame({'A': []})
        return df


file_selector()


FileLocations = pd.DataFrame(FileLocations)
FileLocations = FileLocations.set_index('File Name')
FileLocations = FileLocations.drop_duplicates()

shelf_files = 'shelve.out'
my_shelf = shelve.open(shelf_files)

# TODO: Make filename dependent
Year = datetime.datetime.now().year
Month = datetime.datetime.now().month

Year = 2020
Month = 1

try:
    filename = my_shelf['current_file']
    filename = 'January Report 2020.xlsx'
    wb = load_workbook(filename)
    print('Using the existing file')
    work_sheet = wb.active
except FileNotFoundError:   # This error is because though a filename is on the shelf, but it isn't in the folder
    # TODO: Ask for a new filename
    filename = my_shelf['current_file']
    print('%s was the last one used, but can\'t be found, please end this program and '
          'save it in the folder, or create a new file name' % filename)
    filename = input() + '.xlsx'
    my_shelf['current_file'] = filename
    try:
        wb = load_workbook(filename)
        print('Found the file')
    except FileNotFoundError:
        wb = Workbook()
        print('Creating a new file')

    work_sheet = wb.active
    work_sheet.title = calendar.month_name[Month] + ' 2020'  # TODO: make filename dependent
    my_shelf['SheetName'] = work_sheet.title
except KeyError:    # This means that there is no filename on the shelf
    # TODO: Ask for a new filename
    filename = input('What is the filename you would like to use?') + '.xlsx'
    my_shelf['current_file'] = filename
    try:
        wb = load_workbook(filename)
        print('Found the file')
    except FileNotFoundError:
        wb = Workbook()
        print('Creating a new file')
    work_sheet = wb.active
    work_sheet.title = calendar.month_name[Month] + ' 2020'  # TODO: make filename dependent
    my_shelf['SheetName'] = work_sheet.title

if FileLocations['Location']['Directory']:
    for index, folder_files in enumerate(os.listdir(FileLocations['Location']['Directory'])):
        for report in range(len(files)):
            if re.search(files[report], folder_files, re.IGNORECASE):
                if files[report] == 'Tender':
                    TenderReport = FileLocations['Location']['Directory'] + '/' + folder_files
                    while True:
                        try:
                            TenderedHigherNames = pd.read_excel(TenderReport, skiprows=4).columns
                            Tendered = pd.read_excel(TenderReport, skiprows=7)
                            Tendered = Tendered.set_index(['Unnamed: 0'])
                            Tendered.index = pd.Series(Tendered.index).fillna(method='ffill')
                            break
                        except NameError:
                            print("The Tender Report must be named 'Tender' only")
                            TenderReport = file_selector()  # TODO: Get file specifically

                if files[report] == 'Employee Sale':
                    EmpSales = FileLocations['Location']['Directory'] + '/' + folder_files

                    while True:
                        try:
                            EmpDisc = pd.read_excel(EmpSales)
                            EmpDisc = EmpDisc.set_index(['Store Name'])
                            break
                        except NameError:
                            print("The Employee Sales Report must be named 'Employee Sale' only")
                            EmpSales = file_selector() #TODO: Get EmpSales Specifically

                if files[report] == 'Tax Free Sale':
                    TaxFreeSales = FileLocations['Location']['Directory'] + '/' + folder_files

                    while True:
                        try:
                            Tax_Exempt = pd.read_excel(TaxFreeSales)
                            Tax_Exempt = Tax_Exempt.set_index(['Store Name'])
                            break
                        except NameError:
                            print("The Tax Exemption Report must be named 'Tax Free Sale'")
                            TaxFreeSales = file_selector() #TODO: Get tax exempted report specifically
                            continue

                if files[report] == 'Purchased GC':
                    GC_Sales = FileLocations['Location']['Directory'] + '/' + folder_files

                    while True:
                        try:
                            PurchasedGC = pd.read_excel(GC_Sales)
                            PurchasedGC = PurchasedGC.set_index('Store Name')
                            break
                        except NameError:
                            print("The Purchased GCs should be named 'Purchased GC' only")
                            GC_Sales = file_selector() #TODO: Get GCSales Location specifically
                            continue

                if files[report] == 'Redeemed GC':
                    GC_Used = FileLocations['Location']['Directory'] + '/' + folder_files

                    while True:
                        try:
                            RedeemedGC = pd.read_excel(GC_Used)
                            RedeemedGC = RedeemedGC.set_index('Store Name')
                            break
                        except NameError:
                            print("The Redeemed GCs should be named 'Redeemed GC' only")
                            GC_Used = file_selector() #TODO: Get GC_Used
                            continue

                if files[report] == 'CM Report':
                    CM_Sales_Issuance = FileLocations['Location']['Directory'] + '/' + folder_files

                    while True:
                        try:
                            CreditMemo = pd.read_excel(CM_Sales_Issuance)
                            break
                        except NameError:
                            print("The Credit Memos should be named 'CM Report' only")
                            CM_Sales_Issuance = file_selector() #TODO: Get CM Sales specifically
                            continue

TaxRate = 'Tax Rate.xlsx'
PurchasedGC = pd.DataFrame({'A' : []})

Tax = pd.read_excel(TaxRate)
Tax = Tax.set_index(['Headquarters'])

CreditMemo['Invoice #'] = CreditMemo['Invoice #'].fillna(0)

try:
    CreditMemo = CreditMemo.astype({'Invoice #': 'int'})
    CreditMemo = CreditMemo.set_index('Invoice #')
except ValueError:
    invoices = []
    for CMs in range(len(CreditMemo)):
        try:
            invoices.append(CreditMemo['Invoice #'][CMs].split()[0])
        except AttributeError:
            invoices.append(CreditMemo['Invoice #'][CMs])
    CreditMemo['Invoice #'] = invoices
    CreditMemo = CreditMemo.set_index('Invoice #')

Locations = ['Alexandria', 'Asheville', 'Austin', 'Baton Rouge', 'Birmingham', 'Boston', 'Buckhead', 'Charleston',
             'Charlotte', 'Chattanooga', 'Chicago', 'Cincinnati', 'Columbia', 'Dallas', 'Detroit', 'Fort Worth',
             'Houston', 'Huntsville', 'Indianapolis', 'Kansas City', 'Knoxville', 'Little Rock', 'Los Angeles',
             'Louisville', 'Marietta', 'Memphis', 'Miami', 'Minneapolis', 'Nashville', 'New Orleans',
             'Orlando', 'Paramus', 'Philadelphia', 'Pittsburgh', 'Portland', 'Raleigh', 'San Antonio',
             'Savannah', 'Sherman Oaks', 'Tampa']

Locations_Key = {
    'Birmingham': 'AL (HOMEWOOD)',
    'Huntsville': 'AL1 (HUNTSVILLE)',
    'Little Rock': 'AR (LITTLE ROCK)',
    'Los Angeles': 'CA2 (LOS ANGELES)',
    # 'Sherman Oaks': 'CA3 (SHERMAN OAKS)',
    'Miami': 'FL (MIAMI)',
    'Orlando': 'FL3 (ORLANDO)',
    'Tampa': 'FL4 (TAMPA)',
    'Marietta': 'GA1 (MARIETTA)',
    'Savannah': 'GA2 (SAVANNAH)',
    'Buckhead': 'GA3 (BUCKHEAD)',
    'Chicago': 'IL (CHICAGO)',
    'Indianapolis': 'IN (INDIANAPOLIS)',
    # 'Louisville': 'KY (LOUISVILLE)',
    'Kansas City': 'KS (KANSAS)',
    'New Orleans': 'LA (NOLA)',
    'Baton Rouge': 'LA1 (BATON ROUGE)',
    # 'Boston': 'MA (BOSTON)',
    'Detroit': 'MI (DETROIT)',
    'Minneapolis': 'MN (MINNEAPOLIS)',
    'Charlotte': 'NC (CHARLOTTE)',
    'Raleigh': 'NC1 (RALEIGH)',
    'Asheville': 'NC2 (ASHEVILLE)',
    'Paramus': 'NJ1 (PARAMUS)',
    'Cincinnati': 'OH (CINCINNATI)',
    'Portland': 'OR1 (WSM)',
    # 'Philadelphia': 'PA (PHILADELPHIA)',
    'Pittsburgh': 'PA1 (PITTSBURGH)',
    'Mount Pleasant': 'SC2 (MTP)',
    'Columbia': 'SC5 (COLUMBIA)',
    'Nashville': 'TN (NASHVILLE)',
    'Memphis': 'TN1 (MEMPHIS)',
    'Knoxville': 'TN2 (KNOXVILLE)',
    'Chattanooga': 'TN3 (CHATTANOOGA)',
    'Dallas': 'TX (DALLAS)',
    'Houston': 'TX1 (HOUSTON)',
    'Austin': 'TX2 (AUSTIN)',
    'San Antonio': 'TX3 (SAN ANTONIO)',
    'Fort Worth': 'TX4 (FT WORTH)',  # Closes near end of January
    'Alexandria': 'VA (ALEXANDRIA)'
}

RED = 'FF0000'
work_sheet.column_dimensions['A'].width = 25
Bold = Font(bold=True)
BoldRed = Font(bold=True, color=RED)
Normal = Font(bold=False)
Hor_Center = Alignment(horizontal='center', vertical='bottom')
Hor_Left = Alignment(horizontal='left', vertical='bottom')
TextWrap = Alignment(wrap_text=True, horizontal='center', vertical='bottom')
TitleBorder = Border(top=Side(border_style='thin', color='000000'),
                     bottom=Side(border_style='thin', color='000000'))


def title(text, working_cell, font, new=False, alignment=None, border=None, merge='A1', number_format='General',
          column_width=False, place=0):
    if new:
        try:
            text = text.iat[place]
        except AttributeError:
            text = text

    work_sheet.merge_cells(merge)
    work_sheet[working_cell].font = font
    work_sheet[working_cell].alignment = alignment
    work_sheet[working_cell] = text
    work_sheet[working_cell].border = border
    work_sheet[working_cell].number_format = number_format
    if column_width:
        COL = ''.join([i for i in working_cell if not i.isdigit()])
        work_sheet.column_dimensions[COL].width = len(work_sheet[working_cell].value) + 2


def data_frame_try_catch(df, group, location, place):
    try:
        search = df[group][location].iat[place]
    except AttributeError:
        search = df[group][location]

    return search


def labels():
    Title = work_sheet['A1']
    Title.alignment = Hor_Center
    Description = work_sheet['A2']
    Description.alignment = Hor_Left
    Title.font = Font(bold=True)
    Description.font = Font(bold=True)

    work_sheet['A1'] = 'Nadeau Corporation'
    work_sheet.merge_cells('A2:B2')
    work_sheet['A2'] = 'Transactions Detail - Month Year'
    work_sheet.column_dimensions['A'].width = 20
    work_sheet.column_dimensions['Z'].width = 20
    work_sheet.column_dimensions['AD'].width = 20
    work_sheet.column_dimensions['B'].width = 20

    work_sheet['A5'] = 'Bank'
    work_sheet['B5'] = 'Date'
    work_sheet['A5'].font = Bold
    work_sheet['B5'].font = Bold

    """
    Start of Row 4 Titles

    """
    title(text='Cash/Check', merge='D4:J4', working_cell='D4', font=Bold, alignment=Hor_Center)

    title(text='CA - Credit Card', merge='L4:R4', working_cell='L4', font=Bold, alignment=Hor_Center)

    title(text='Total Credit Cards', merge='S4:S5', working_cell='S4', font=Bold, alignment=TextWrap, column_width=True)

    title(text='GRAND TOTAL SALES', merge='T4:T5', working_cell='T4', font=Bold, alignment=TextWrap, column_width=True)

    title(text='Total Gift Cards Used', merge='U4:U5', working_cell='U4', font=Bold, alignment=TextWrap,
          column_width=True)

    title(text='TAX INCLUDED', merge='V4:V5', working_cell='V4', font=Bold, alignment=TextWrap, column_width=True)

    title(text='TAX EXEMPT', merge='Y4:Y5', working_cell='Y4', font=Bold, alignment=TextWrap, column_width=True)

    title(text='Total Employee', merge='AA4:AA5', working_cell='AA4', font=Bold, alignment=TextWrap, column_width=True)

    title(text='Total Employee', merge='AE4:AE5', working_cell='AE4', font=Bold, alignment=TextWrap, column_width=True)

    """
    End of Row 4 Titles

    """

    """
    Start of Row 5 Titles

    """
    title(text='Cash', font=Bold, alignment=Hor_Center, merge='D5:F5', border=TitleBorder, working_cell='D5')

    title(text='Check', merge='H5:J5', working_cell='H5', border=TitleBorder, alignment=Hor_Center, font=Bold)

    title(text='Visa/MC/Discover', merge='L5:N5', working_cell='L5', font=Bold, border=TitleBorder,
          alignment=Hor_Center)

    title(text='Amex', merge='P5:R5', working_cell='P5', font=Bold, border=TitleBorder, alignment=Hor_Center)

    title(text='GC bought', working_cell='X5', font=BoldRed, alignment=Hor_Center, column_width=True)

    title(text='PPS2', working_cell='Z5', font=Bold, alignment=Hor_Center)

    title(text='Difference', working_cell='AC5', font=Bold, alignment=Hor_Center, column_width=True)

    title(text='PP1', working_cell='AD5', font=Bold, alignment=Hor_Center)

    title(text='Difference', working_cell='AG5', font=Bold, alignment=Hor_Center, column_width=True)

    """
    End of Row 5 Titles

    """

    subTitles = ['Nadeau Reports', 'Per Statements', 'Difference', ' ']

    Letters = string.ascii_uppercase
    Letters = Letters.replace('ABC', '')
    Letters = Letters.replace('STUVWXYZ', '')

    subs = 0
    for letter in range(len(Letters)):
        title(text=subTitles[subs], working_cell=Letters[letter] + '6', font=Bold, alignment=Hor_Center,
              column_width=True)
        if subs >= 3:
            subs = -1
        subs += 1

    work_sheet.freeze_panes = 'C7'


labels()

RowFillSeparator = PatternFill(fill_type='solid', start_color='FFFF0000', end_color='FFFF0000')
RowBorderSeparator = Border(bottom=Side(style='thick'))

if 'PPL' in TenderedHigherNames:
    Tendered = Tendered.rename(columns={
        'Unnamed: 7': 'Date',
        'AMT': 'Cash', 'INV_TAXABLE_TOTAL': 'Cash Commission', 'INV_EXT_LINE_TAX_AMT': 'Cash Taxed',
        'AMT.1': 'Check', 'INV_TAXABLE_TOTAL.1': 'Check Commission', 'INV_EXT_LINE_TAX_AMT.1': 'Check Taxed',
        'AMT.2': 'AMEX', 'INV_TAXABLE_TOTAL.2': 'AMEX Commission', 'INV_EXT_LINE_TAX_AMT.2': 'AMEX Taxed',
        'AMT.3': 'VisaMCD', 'INV_TAXABLE_TOTAL.3': 'VisaMCD Commission', 'INV_EXT_LINE_TAX_AMT.3': 'VisaMCD Taxed',
        'AMT.4': 'CCTotal', 'INV_TAXABLE_TOTAL.4': 'CCTotal Commission', 'INV_EXT_LINE_TAX_AMT.4': 'CCTotal Taxed',
        'AMT.5': 'GCTotal', 'INV_TAXABLE_TOTAL.5': 'GCTotal Commission', 'INV_EXT_LINE_TAX_AMT.5': 'GCTotal Taxed',
        'AMT.7': 'SCTotal', 'INV_TAXABLE_TOTAL.7': 'SCTotal Commission', 'INV_EXT_LINE_TAX_AMT.7': 'SCTotal Taxed',
        'AMT.8': 'GTotal', 'INV_TAXABLE_TOTAL.8': 'GTotal Commission', 'INV_EXT_LINE_TAX_AMT.8': 'GTotal Taxed'})
elif 'Check' not in TenderedHigherNames:
    Tendered = Tendered.rename(columns={
        'Unnamed: 7': 'Date',
        'AMT': 'Cash', 'INV_TAXABLE_TOTAL': 'Cash Commission', 'INV_EXT_LINE_TAX_AMT': 'Cash Taxed',
        #'AMT.1': 'Check', 'INV_TAXABLE_TOTAL.1': 'Check Commission', 'INV_EXT_LINE_TAX_AMT.1': 'Check Taxed',
        'AMT.1': 'AMEX', 'INV_TAXABLE_TOTAL.1': 'AMEX Commission', 'INV_EXT_LINE_TAX_AMT.1': 'AMEX Taxed',
        'AMT.2': 'VisaMCD', 'INV_TAXABLE_TOTAL.2': 'VisaMCD Commission', 'INV_EXT_LINE_TAX_AMT.2': 'VisaMCD Taxed',
        'AMT.3': 'CCTotal', 'INV_TAXABLE_TOTAL.3': 'CCTotal Commission', 'INV_EXT_LINE_TAX_AMT.3': 'CCTotal Taxed',
        'AMT.4': 'GCTotal', 'INV_TAXABLE_TOTAL.4': 'GCTotal Commission', 'INV_EXT_LINE_TAX_AMT.4': 'GCTotal Taxed',
        'AMT.5': 'SCTotal', 'INV_TAXABLE_TOTAL.5': 'SCTotal Commission', 'INV_EXT_LINE_TAX_AMT.5': 'SCTotal Taxed',
        'AMT.6': 'GTotal', 'INV_TAXABLE_TOTAL.6': 'GTotal Commission', 'INV_EXT_LINE_TAX_AMT.6': 'GTotal Taxed'})

else:
    Tendered = Tendered.rename(columns={
        'Unnamed: 7': 'Date',
        'AMT': 'Cash', 'INV_TAXABLE_TOTAL': 'Cash Commission', 'INV_EXT_LINE_TAX_AMT': 'Cash Taxed',
        'AMT.1': 'Check', 'INV_TAXABLE_TOTAL.1': 'Check Commission', 'INV_EXT_LINE_TAX_AMT.1': 'Check Taxed',
        'AMT.2': 'AMEX', 'INV_TAXABLE_TOTAL.2': 'AMEX Commission', 'INV_EXT_LINE_TAX_AMT.2': 'AMEX Taxed',
        'AMT.3': 'VisaMCD', 'INV_TAXABLE_TOTAL.3': 'VisaMCD Commission', 'INV_EXT_LINE_TAX_AMT.3': 'VisaMCD Taxed',
        'AMT.4': 'CCTotal', 'INV_TAXABLE_TOTAL.4': 'CCTotal Commission', 'INV_EXT_LINE_TAX_AMT.4': 'CCTotal Taxed',
        'AMT.5': 'GCTotal', 'INV_TAXABLE_TOTAL.5': 'GCTotal Commission', 'INV_EXT_LINE_TAX_AMT.5': 'GCTotal Taxed',
        'AMT.6': 'SCTotal', 'INV_TAXABLE_TOTAL.6': 'SCTotal Commission', 'INV_EXT_LINE_TAX_AMT.6': 'SCTotal Taxed',
        'AMT.7': 'GTotal', 'INV_TAXABLE_TOTAL.7': 'GTotal Commission', 'INV_EXT_LINE_TAX_AMT.7': 'GTotal Taxed'})

Currency = '$#,##0.00_);[Red]($#,##0.00)'

Month_Range = calendar.monthrange(Year, Month)
Locations_Info = []
CellValue = 6
MonthFirst = 0
for BankIndex, Bank in enumerate(Locations_Key.keys()):
    for FullMonth in range(Month_Range[1]):

        CurrentTax = Tax['tax'][Bank.replace('\n', '')]
        Date = datetime.date(Year, Month, (FullMonth + 1))
        CellValue += 1
        Row = str(CellValue)
        title(text=Locations_Key[str(Bank)], working_cell=('A' + Row), font=Bold)  # BANK
        title(text=Date.strftime('%m/%d/%Y'), working_cell=('B' + Row), font=Bold)  # DATE
        title(text=('=E%s-D%s' % (Row, Row)), working_cell=('F%s' % Row), font=Normal,
              number_format=Currency)
        title(text=('=I%s-H%s' % (Row, Row)), working_cell=('J%s' % Row), font=Normal,
              number_format=Currency)
        title(text=('=M%s-L%s' % (Row, Row)), working_cell=('N%s' % Row), font=Normal,
              number_format=Currency)
        title(text=('=Q%s-P%s' % (Row, Row)), working_cell=('R%s' % Row), font=Normal,
              number_format=Currency)
        title(text=('=V%s/(T%s-V%s+X%s-Y%s)' % (Row, Row, Row, Row, Row)), working_cell=('W%s' % Row), font=Normal,
              number_format='0.0000')
        title(text=('=L%s+P%s' % (Row, Row)), working_cell='S' + Row, font=Normal,
              number_format=Currency)
        title('=S%s+D%s+H%s+U%s' % (Row, Row, Row, Row), working_cell='T' + Row, font=Normal,
              number_format=Currency)
        if Date.day <= 15:
            title(text=('=T%s-V%s-AE%s+X%s' % (Row, Row, Row, Row)), working_cell='AD' + Row,
                  font=Normal, number_format=Currency)
        else:
            title(text=('=T%s-V%s-AE%s+X%s' % (Row, Row, Row, Row)), working_cell='Z' + Row,
                  font=Normal, number_format=Currency)

        UnTaxed = 0
        GCsPurchased = 0
        GCsRedeemed = 0
        if Bank in Tendered.index:
            try:
                len(PurchasedGC['Date'][Bank])
                PurGo = True
                for GCs in enumerate(pd.to_datetime(PurchasedGC['Date'][Bank])):
                    if Date == GCs[1]:
                        GCsPurchased -= PurchasedGC['GIVEN'][Bank].iat[GCs[0]]
                        title(text=GCsPurchased, working_cell='X' + Row, font=Normal, number_format=Currency)
            except TypeError:
                PurGo = False
                if pd.to_datetime(PurchasedGC['Date'][Bank]) == Date:
                    GCsPurchased = -PurchasedGC['GIVEN'][Bank]
                    title(text=GCsPurchased, working_cell='X' + Row, font=Normal, number_format=Currency)
            except KeyError:
                pass

            try:
                len(RedeemedGC['Date'][Bank])
                RedGo = True
                for RGC in enumerate(pd.to_datetime(RedeemedGC['Date'][Bank])):
                    if Date == RGC[1]:
                        GCsRedeemed += RedeemedGC['TAKEN'][Bank].iat[RGC[0]]
                        title(text=GCsRedeemed, working_cell='U' + Row, font=Normal, number_format=Currency)
            except TypeError:
                RedGo = False
                if Date == pd.to_datetime(RedeemedGC['Date'][Bank]):
                    GCsRedeemed = RedeemedGC['TAKEN'][Bank]
                    title(text=GCsRedeemed, working_cell='U' + Row, font=Normal, number_format=Currency)
            except KeyError:
                pass

            for data in range(len(Tendered['Date'][Bank])):
                try:
                    working_date = pd.to_datetime(Tendered['Date'][Bank].iat[data])
                except AttributeError:
                    working_date = pd.to_datetime(Tendered['Date'][Bank])
                if Date == working_date:
                    try:
                        UnTaxed = Memo = 0
                        for TaxFree in range(len(Tax_Exempt['Date'][Bank])):
                            if Date == pd.to_datetime(Tax_Exempt['Date'][Bank].iat[TaxFree]):
                                UnTaxed += Tax_Exempt['Item Subtotal'][Bank].iat[TaxFree]
                                try:
                                    Memo += CreditMemo['AMT'][str(Tax_Exempt['Invoice #'][Bank].iat[TaxFree])]
                                except (KeyError, IndexError):
                                    pass
                        title(text=UnTaxed - Memo, working_cell=('Y%s' % Row), font=Normal, number_format=Currency)
                        UnTaxed = Memo = 0
                    except KeyError:
                        pass
                    except TypeError:
                        """
                        Sometimes there is only one Credit Memo, so we can't use .iat function to count through the CMs
                        """
                        TaxFree = 0
                        if Date == pd.to_datetime(Tax_Exempt['Date'][Bank]):
                            title(text=(Tax_Exempt['Item Subtotal'][Bank]), working_cell=('Y%s' % Row), font=Normal,
                                  number_format=Currency)

                    # Cash Total
                    title(text=Tendered['Cash'][Bank], working_cell=('D' + Row), font=Normal,
                          number_format=Currency, new=True, place=data)
                    # Check Total
                    #title(text=Tendered['Check'][Bank], working_cell='H' + Row, font=Normal,
                     #     number_format=Currency, new=True, place=data)
                    title(text=Tendered['VisaMCD'][Bank], working_cell='L' + Row, font=Normal,
                          number_format=Currency, new=True, place=data)
                    title(text=Tendered['AMEX'][Bank], working_cell='P' + Row, font=Normal,
                          number_format=Currency, new=True, place=data)

                    GCsPurchased = 0
                    GCsRedeemed = 0
                    RedGo = False
                    PurGo = False

                    # ****SC is purposefully being removed from the GTotal & GTotal Taxed****

                    GrandTotal = data_frame_try_catch(df=Tendered, group='GTotal', location=Bank, place=data)

                    GCTotal = data_frame_try_catch(df=Tendered, group='GCTotal', location=Bank, place=data)

                    if not np.isnan(GCTotal):
                        if GCTotal < 0:
                            GrandTotal -= GCTotal

                    SCTotal = data_frame_try_catch(df=Tendered, group='SCTotal', location=Bank, place=data)

                    if np.isnan(SCTotal):
                        GTTaxed = data_frame_try_catch(df=Tendered, group='GTotal Taxed', location=Bank, place=data)
                        title(text=round(GTTaxed, 2), working_cell='V' + Row,
                              font=Normal, number_format=Currency)
                        if Date.day <= 15:
                            # PAY PERIOD
                            # This Try Block is adding the Employee payments
                            try:
                                try:
                                    EmpDayTotal = 0
                                    for discount in range(len(EmpDisc['Date'][Bank])):

                                        EmpDiscDate = data_frame_try_catch(df=EmpDisc, group='Date',
                                                                           location=Bank, place=discount)

                                        if Date.day == pd.to_datetime(EmpDiscDate).day:

                                            EmpDayTotal += data_frame_try_catch(df=EmpDisc, group='Item Subtotal',
                                                                                location=Bank, place=discount)
                                            title(text=EmpDayTotal, working_cell='AE' + Row, font=Normal,
                                                  number_format=Currency)
                                except TypeError:
                                    if Date.day == pd.to_datetime(EmpDisc['Date'][Bank]).day:
                                        title(text=EmpDisc['Item Subtotal'][Bank], working_cell='AE' + Row, font=Normal,
                                              number_format=Currency)
                                    #  print(EmpDisc['Item Subtotal'][Bank], Date.day)
                            except KeyError:
                                pass
                        else:
                            # PAY PERIODS
                            title(text=('=T%s-V%s-AA%s+X%s' % (Row, Row, Row, Row)), working_cell='Z' + Row,
                                  font=Normal, number_format=Currency)
                            # This Try Block adds the Employee payments
                            try:
                                try:
                                    EmpDayTotal = 0
                                    for discount in range(len(EmpDisc['Date'][Bank])):
                                        if Date == pd.to_datetime(EmpDisc['Date'][Bank].iat[discount]):
                                            EmpDayTotal += EmpDisc['Item Subtotal'][Bank].iat[discount]
                                        title(text=EmpDayTotal, working_cell='AA' + Row, font=Normal,
                                              number_format=Currency)
                                except TypeError:
                                    title(text=EmpDisc['Item Subtotal'][Bank], working_cell='AA' + Row, font=Normal,
                                          number_format=Currency)
                                    break
                            except KeyError:
                                pass
                    else:
                        # PAY PERIODS
                        elseSC = data_frame_try_catch(df=Tendered, group='SCTotal', location=Bank, place=data)
                        SC_COM = round(elseSC / (CurrentTax + 1), 2)
                        SCTax = round((elseSC - SC_COM), 2)
                        # Tax Calculations
                        elseSCTaxed = data_frame_try_catch(df=Tendered,
                                                           group='SCTotal Taxed', location=Bank, place=data)
                        if np.isnan(elseSCTaxed):
                            elseGCTaxed = data_frame_try_catch(df=Tendered,
                                                                   group='GTotal Taxed', location=Bank, place=data)
                            if elseSC < 0:

                                title(text=(round(elseGCTaxed, 2)
                                            - SCTax), working_cell='V' + Row, font=Normal, number_format=Currency)
                            else:  # Pretty sure these do the same thing...
                                title(text=(round(elseGCTaxed, 2)
                                            - SCTax), working_cell='V' + Row, font=Normal, number_format=Currency)

                        else:
                            elseGTTaxed = data_frame_try_catch(df=Tendered, group='GTotal Taxed',
                                                               location=Bank, place=data)
                            if elseSC < 0:
                                title(text=(round(elseGTTaxed, 2)
                                            - SCTax), working_cell='V' + Row, font=Normal, number_format=Currency)
                            else:
                                # SCTax += abs(Tendered['SCTotal Taxed'][Bank].iat[data])
                                title(text=(round(elseGTTaxed, 2)
                                            - SCTax), working_cell='V' + Row, font=Normal, number_format=Currency)

                        if Date.day <= 15:
                            # PAY PERIODS
                            # This Try Block adds the Employee payments
                            try:
                                try:
                                    EmpDayTotal = 0
                                    for discount in range(len(EmpDisc['Date'][Bank])):
                                        fifEmpDate = data_frame_try_catch(df=EmpDisc, group='Date',
                                                                          location=Bank, place=discount)
                                        if Date.day == pd.to_datetime(fifEmpDate).day:

                                            EmpDayTotal += data_frame_try_catch(df=EmpDisc, group='Item Subtotal',
                                                                               location=Bank, place=discount)
                                            title(text=EmpDayTotal, working_cell='AE' + Row, font=Normal,
                                                  number_format=Currency)
                                except TypeError:
                                    if Date.day == pd.to_datetime(EmpDisc['Date'][Bank]).day:
                                        title(text=(EmpDisc['Item Subtotal'][Bank]), working_cell='AE' + Row,
                                              font=Normal,
                                              number_format=Currency)
                                    break
                            except KeyError:
                                pass
                        else:
                            title(text=('=T%s-V%s-AA%s+X%s' % (Row, Row, Row, Row)), working_cell='Z' + Row,
                                  font=Normal, number_format=Currency)
                            # This Try Block adds the Employee payments
                            try:
                                try:
                                    EmpDayTotal = 0
                                    for discount in range(len(EmpDisc['Date'][Bank])):
                                        if Date == pd.to_datetime(EmpDisc['Date'][Bank].iat[discount]):
                                            EmpDayTotal += EmpDisc['Item Subtotal'][Bank].iat[discount]
                                    title(text=EmpDayTotal, working_cell='AA' + Row, font=Normal,
                                          number_format=Currency)
                                except TypeError:
                                    title(text=EmpDisc['Item Subtotal'][Bank], working_cell='AA' + Row, font=Normal,
                                          number_format=Currency)
                            except KeyError:
                                pass
    # Full month's total calculations
    CellValue += 1
    title(text=(Locations_Key[str(Bank)].split(' ')[0] + ' Total'), working_cell=('A' + str(CellValue)), font=Bold,
          number_format=Currency)  # Total

    title(text=('=SUM(D%s:D%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('D%s' % CellValue), font=Normal,
          number_format=Currency)  # Cash
    title(text=('=SUM(E%s:E%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('E%s' % CellValue), font=Normal,
          number_format=Currency)
    title(text=('=E%s-D%s' % (CellValue, CellValue)), working_cell=('F%s' % CellValue), font=Normal,
          number_format=Currency)

    title(text=('=SUM(H%s:H%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('H%s' % CellValue), font=Normal,
          number_format=Currency)  # Check
    title(text=('=SUM(I%s:I%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('I%s' % CellValue), font=Normal,
          number_format=Currency)
    title(text=('=I%s-H%s' % (CellValue, CellValue)), working_cell=('J%s' % CellValue), font=Normal,
          number_format=Currency)

    title(text=('=SUM(L%s:L%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('L%s' % CellValue), font=Normal,
          number_format=Currency)  # Visa
    title(text=('=SUM(M%s:M%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('M%s' % CellValue), font=Normal,
          number_format=Currency)
    title(text=('=M%s-L%s' % (CellValue, CellValue)), working_cell=('N%s' % CellValue), font=Normal,
          number_format=Currency)

    title(text=('=SUM(P%s:P%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('P%s' % CellValue), font=Normal,
          number_format=Currency)  # AMEX
    title(text=('=SUM(Q%s:Q%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('Q%s' % CellValue), font=Normal,
          number_format=Currency)
    title(text=('=Q%s-P%s' % (CellValue, CellValue)), working_cell=('R%s' % CellValue), font=Normal,
          number_format=Currency)

    title(text=('=SUM(S%s:S%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('S%s' % CellValue), font=Normal,
          number_format=Currency)
    title(text=('=SUM(T%s:T%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('T%s' % CellValue), font=Normal,
          number_format=Currency)
    title(text=('=SUM(U%s:U%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('U%s' % CellValue), font=Normal,
          number_format=Currency)
    title(text=('=SUM(V%s:V%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('V%s' % CellValue), font=Normal,
          number_format=Currency)

    title(text=('=SUM(X%s:X%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('X%s' % CellValue), font=Normal,
          number_format=Currency)
    title(text=('=SUM(Z%s:Z%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('Z%s' % CellValue), font=Normal,
          number_format=Currency)
    title(text=('=SUM(AA%s:AA%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('AA%s' % CellValue),
          font=Normal, number_format=Currency)

    title(text=('=SUM(AD%s:AD%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('AD%s' % CellValue),
          font=Normal, number_format=Currency)
    title(text=('=SUM(AE%s:AE%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('AE%s' % CellValue),
          font=Normal, number_format=Currency)
    title(text=('=SUM(Z%s:AD%s)' % (int(Row) - FullMonth, CellValue - 1)), working_cell=('AG%s' % CellValue),
          font=Normal, number_format=Currency)

    RED = 'AA0000'
    babyBLUE = '00ABFF'
    YELLOW = 'FFFF00'
    # CurrentTax = globals().get('CurrentTax')
    if CurrentTax > 0:
        rule = ColorScaleRule(start_type='num', start_value=CurrentTax - 0.01, start_color=RED,
                              mid_type='num', mid_value=CurrentTax, mid_color=babyBLUE,
                              end_type='num', end_value=CurrentTax + 0.01, end_color=YELLOW)
    else:
        rule = ColorScaleRule(start_type='num', start_value=0.0000, start_color=babyBLUE,
                              mid_type='num', mid_value=CurrentTax, mid_color=babyBLUE,
                              end_type='num', end_value=CurrentTax + 0.01, end_color=YELLOW)

    # dxf = DifferentialStyle(fill=RED)

    work_sheet.conditional_formatting.add(('W%s:W%s' % (int(Row) - FullMonth, CellValue - 1)), rule)

    for cell in work_sheet[CellValue:CellValue]:
        cell.border = RowBorderSeparator

    CellValue += 1

if __name__ == '__main__':
    worksheet.dimensions.ColumnDimension(work_sheet, bestFit=True)
    wb.save(filename=filename)
    """
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
    """
