from django.shortcuts import render,redirect
import xlrd 
import requests
import json
import xlwt 
from xlwt import Workbook
from django.http import HttpResponse
from django.contrib import messages
from django.utils.encoding import smart_str
from openpyxl import load_workbook
from pandas import DataFrame


def index(request):
    if "GET" == request.method:
        return render(request, 'geoCodeXML/index.html')
    else:
        filePathName = request.POST.get('pathOfXLFile')
        XLworkBook = xlrd.open_workbook(filePathName)
        sheet = XLworkBook.sheet_by_index(0)

        sheet.cell_value(0, 0)
        
        wb = Workbook()
        finalResult = wb.add_sheet('finalResult')
        finalResult.write(0, 0, 'Addresses')
        finalResult.write(0, 1, 'lattitude')
        finalResult.write(0, 2, 'longitude')


        googleGeoCodingAPIKey = 'AIzaSyDscGlVbYHZvb5vO4GHFoucYlqXF5J_Zc8'
        for i in range(1,sheet.nrows):
            address = sheet.cell_value(i, 0)
            addressList = address.split(" ")
            temp = ''
            for data in address:
                temp +=data
            requestURL = 'https://maps.googleapis.com/maps/api/geocode/json?address='+temp+'&key='+googleGeoCodingAPIKey
            APIresponse = requests.get(requestURL)
            APIresponseJSON = APIresponse.json()
            finalResult.write(i,0,address)
            finalResult.write(i,1,APIresponseJSON.get("results")[0].get("geometry").get("location").get("lat"))
            finalResult.write(i,2,APIresponseJSON.get("results")[0].get("geometry").get("location").get("lng"))
            wb.save('finalResultXL.xls')
        f = open('finalResult.xls','rb')
        response = HttpResponse(f, content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename="finalResult.xls"'
        messages.success(request,'File Processed Successfully :)')
        return response
