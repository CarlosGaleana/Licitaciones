#-----------------------------------------------------
# Intel Corporation - SMG SS LATAM E&G
# Web Scrapping from Colombia Compra Eficiente POs
# 
# This code is an adaptation from the original version
# made by Nelson Rojas.
#
# Editor: Kimberly Orozco-Retana ITS.Intern
#-----------------------------------------------------

import pandas as pd
from bs4 import BeautifulSoup
import requests
import xlsxwriter

# Paste path of downloaded file from Colombia Compra Eficiente given desired dates
path = r"C:\Users\kdorozco\OneDrive - Intel Corporation\Desktop\Automatization_Projects\colombia\data\ETP03_01JUL24SEP.xls"
# Name the output file for scraped data
xlsx_file = 'Items_ETP3.xlsx'

def get_items_PO(PO_Path,file_name):
    data = pd.read_excel(PO_Path)
    orders = data.Orden

    url_base = 'https://www.colombiacompra.gov.co/tienda-virtual-del-estado-colombiano/ordenes-compra/'
    PO_info = {}
    d = 1
    for i in orders:
        per = str(round(100*d/len(orders),2))+'%' #shows % progress in scrapping
        PO_info[i] = {}                                  
        url = url_base + str(i)
        page = requests.get(url).text
        soup = BeautifulSoup(page,'lxml')

        #Find all table rows in the soup
        table_rows = soup.find_all("tr")
        
        for row in table_rows:
            #Extract the text from all table cells in this row
            cells = row.find_all("td")
            cell_text = [cell.get_text() for cell in cells]

            #Check if the row contains data (e.g., has 6 cells)
            if len(cell_text) == 6:
                key = cell_text[0]
                PO_info[i][key] ={
                    'articulo':cell_text[1],
                    'cantidad':cell_text[2],
                    'precio':cell_text[4].replace(".","").replace(',','.'),
                    'total':cell_text[5].replace(".","").replace(',','.')
                }
        
        print(per)
        d += 1

    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet('Artículos')
    worksheet.write(0,0,'Orden')
    worksheet.write(0,1,'Entidad estatal')
    worksheet.write(0,2,'Fecha')
    worksheet.write(0,3,'Artículo')
    worksheet.write(0,4,'Proveedor')
    worksheet.write(0,5,'Unidades')
    worksheet.write(0,6,'Precio por Unidad')
    worksheet.write(0,7,'Total')
    worksheet.write(0,8,'Descripción')

    k = 1
    for codigo in PO_info:
        for no in PO_info[codigo]:
            worksheet.write(k, 0, str(codigo))
            worksheet.write(k, 3, PO_info[codigo][no]['articulo'])
            worksheet.write(k, 5, PO_info[codigo][no]['cantidad'])
            worksheet.write(k, 6, PO_info[codigo][no]['precio'].strip("."))
            worksheet.write(k, 7, PO_info[codigo][no]['total'].strip("."))
            k += 1
            
    workbook.close()

get_items_PO(path, xlsx_file)