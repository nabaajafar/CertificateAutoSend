# Python 3 code to rename multiple 
# files in a directory or folder

# importing os module
import os
import os, os.path
# data from excel
import xlwings as xw
# import openpyxl
# import pandas as pd


# Function to rename multiple files
def main():
    # Specifying a sheet
    ws = xw.Book("C:\\Users\\HUAWEI\\Desktop\\abc.xlsx").sheets['Sheet1']
    # ws = openpyxl.load_workbook("abc.xlsx")
    # ws = pd.read_excel('abc.xlsx')
    # Selecting data from
    # a single cell
    v1 = ws.range("A2:A24").value
    file_name_base = "KFU" # the pdf name start with KFU {changed number}
    # folder = "C:\Users\HUAWEI\Desktop\cert"
    folder = "C:\\Users\\HUAWEI\\Desktop\\cert"
    # for count, filename in enumerate(os.listdir(folder)):
    # for count, newfilename in range(v1):
    for count, newfilename in enumerate(v1):
        # dst = f"{str(v1[count])}.pdf"
        # src =f"{folder}/{filename}" # foldername/filename, if .py file is outside folder
        # dst =f"{folder}/{dst}"
        print(count)
        src = f"{folder}/{file_name_base} {count+1}.pdf"
        dst = f"{folder}/{newfilename}.pdf"
		# rename() function will
		# rename all the files
        os.rename(src, dst)

# Driver Code
if __name__ == '__main__':
	
	# Calling main() function
	main()



 
