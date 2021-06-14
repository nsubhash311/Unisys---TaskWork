import cv2
import pytesseract
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
import xlsxwriter

#Reading image
path = "D:\OCR\license.jpg"
image = cv2.imread(path)
# print(image)

#After reading the file, we will extract infromation about the license owner from the image.

#Extracting Name
crop1 = image[98:117,175:350]
crop1 = cv2.cvtColor(crop1, cv2.COLOR_BGR2GRAY)
ret, crop1 = cv2.threshold(crop1, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
cv2.imwrite("Pro1.jpg",crop1)


#Extracting  Gaurdian name
crop2 = image[114:133,177:350]
crop2 = cv2.cvtColor(crop2, cv2.COLOR_BGR2GRAY)
ret, crop2 = cv2.threshold(crop2, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
cv2.imwrite("Pro2.jpg",crop2)


#Extracting Address
crop3 = image[165:202,175:500]
crop3 = cv2.cvtColor(crop3, cv2.COLOR_BGR2GRAY)
ret, crop3 = cv2.threshold(crop3, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
cv2.imwrite("Pro3.jpg",crop3)


#Extracting DOB
crop4 = image[132:152,235:390]
crop4 = cv2.cvtColor(crop4, cv2.COLOR_BGR2GRAY)
ret, crop4 = cv2.threshold(crop4, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
cv2.imwrite("Pro4.jpg",crop4)



#Extracting DL number
crop5 = image[80:99,40:378]
crop5 = cv2.cvtColor(crop5, cv2.COLOR_BGR2GRAY)
ret, crop5 = cv2.threshold(crop5, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
cv2.imwrite("Pro5.jpg",crop5)



# Using pytesseract to extract text from all croped images
pytesseract.pytesseract.tesseract_cmd = "C:\Program Files\Tesseract-OCR\\tesseract.exe"
result1 = pytesseract.image_to_string(crop1)
result2 = pytesseract.image_to_string(crop2)
result3 = pytesseract.image_to_string(crop3)
result4 = pytesseract.image_to_string(crop4)
result5 = pytesseract.image_to_string(crop5)

#Some processing
result5 = result5.split(" ")
result3 = result3.split("\n")
k = ""
for st in result3:
	k = k+st
	k = k+" "

result3 = k 

#Printing results on screen
print("Name of card owner :",result1)
print("S/D/W of :",result2)
print("Address :",result3)
print("Date of Birth :",result4)
print("DL number :",result5[-1])


# Writing data to excel file

#create file (workbook) and worksheet
outWorkbook = xlsxwriter.Workbook("data.xlsx")
outSheet = outWorkbook.add_worksheet()

#declare data
values = [result1, result2, result3, result4, result5[-1] ]

#write headers
outSheet.write("A1","Name")
outSheet.write("B1","S/D/W of")
outSheet.write("C1","Address")
outSheet.write("D1","Date of Birth")
outSheet.write("E1","DL number")



#write data to file
outSheet.write("A2",values[0])
outSheet.write("B2",values[1])
outSheet.write("C2",values[2])
outSheet.write("D2",values[3])
outSheet.write("E2",values[4])




outWorkbook.close()
print("Details have been recorded to an excel file.")


# convert to json
excel_data_df = pd.read_excel('data.xlsx')

json_str = excel_data_df.to_json()

print('Excel Sheet to JSON:\n', json_str)