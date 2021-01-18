import os
import pandas as pd
from zipfile import ZipFile
# we use xlsxwriter, openpyxl as the engine for pd.ExcelWriter
from datetime import datetime
import concurrent.futures


startTime = datetime.now()

# replace with path of folder containing sheets that are to have module appended
sheetsLocation = r'C:\Users\Kareem Sayed\Documents\Testr'
# replace with path of an xlsm file containing the module
moduleLocation = r'C:\Users\Kareem Sayed\Documents\Testr\Source.xlsm'


def vbaExtract(moduleLoc):  # extracts the module from the source .xlsm file
    vba_filename = 'vbaProject.bin'
    xlsm_file = moduleLoc
    xlsm_zip = ZipFile(xlsm_file, 'r')

    # Read the xl/vbaProject.bin file.
    vba_data = xlsm_zip.read('xl/' + vba_filename)

    # Write the vba data to a local file.
    vba_file = open(vba_filename, "wb")
    vba_file.write(vba_data)
    vba_file.close()


# extract module from source
vbaExtract(moduleLocation)





def appendMacro(filename):
    # no need to rewrite source document
    if filename.name.endswith('.xlsm') and filename.name != 'Source.xlsm':
        # store data in current workbook, must use openpyxl, xlrd does not support .xlsm
        ws_dict = pd.read_excel(filename, sheet_name=None, engine='openpyxl')
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            # appends macro to *filename*
            writer.book.add_vba_project('./vbaProject.bin')
            for ws_name, df_sheet in ws_dict.items():
                # must set index and header to false else every cell value is shifted from its original value
                df_sheet.to_excel(writer, sheet_name=ws_name, index=False, header=False)


def thread():
    if __name__ == '__main__':
        with concurrent.futures.ThreadPoolExecutor(max_workers=12) as executor:
            executor.map(appendMacro, os.scandir(sheetsLocation))

thread()
print("Total time taken:")
print(datetime.now() - startTime)
