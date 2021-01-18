# Packages Requirements
We use xlsxwriter, openpyxl although we do not import them. They are used as engines for Pandas. Zipfile is essential to export VBAProject.bin from the Source .xlsm document. Multithreading is used to reduce cost of I/O. Currently looking for ways to bypass individual file I/O.

# Next Steps
productionize by letting users input their own path to a VBA bin and excel file.
