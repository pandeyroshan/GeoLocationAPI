from django.shortcuts import render
import xlrd 
import requests
import json
import xlwt 
from xlwt import Workbook
from django.http import HttpResponse

def index(request):
    if "GET" == request.method:
        return render(request, 'geoCodeXML/index.html', {})
    else:
        uploaded_file_path = request.POST.get('pathOfXLFile')
        wb = xlrd.open_workbook(uploaded_file_path)
        sheet = wb.sheet_by_index(0)
        sheet.cell_value(0, 0)
        wb = Workbook()
        sheet1 = wb.add_sheet('Sheet1')
        sheet1.write(0, 0, 'Addresses')
        sheet1.write(0, 1, 'lattitude')
        sheet1.write(0, 2, 'longitude')
        googleGeoCodingAPIKey = 'AIzaSyD4YHhE8VsMnWgFPV0NXy_FQCuzkw-OAjo'
        for i in range(1,sheet.nrows):
            address = sheet.cell_value(i, 0)
            addressList = address.split(" ")
            temp = ''
            for data in address:
                temp +=data
            requestURL = 'https://maps.googleapis.com/maps/api/geocode/json?address='+temp+'&key='+googleGeoCodingAPIKey
            APIresponse = requests.get(requestURL)
            dataDictionary = APIresponse.json()
            sheet1.write(i,0,address)
            sheet1.write(i,1,dataDictionary.get("results")[0].get("geometry").get("location").get("lat"))
            sheet1.write(i,2,dataDictionary.get("results")[0].get("geometry").get("location").get("lng"))
            wb.save('finalResult.xls')
        response = HttpResponse(wb, content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename="finalResult.xls"'
        return response