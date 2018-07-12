"""
A sample for working with the Pyvot library for Excel.
"""

from os import path
from openpyxl import load_workbook
from Service import Service
import services_to_excel

def main():
    """Loads and modifies a sample spreadsheet in Excel."""
    workbook_path = path.join(path.dirname(__file__), 'SILLA DE RUEDAS.xlsx')
    workbook = load_workbook(workbook_path, read_only=True)

    sheet_names = workbook.sheetnames

    for sheet in sheet_names:
        process_sheet(sheet, workbook)
    


services = []
def process_sheet(sheet, workbook):   

    if not is_valid_sheet_name(sheet):
        return

    sheet_ranges = workbook[sheet]    

    local = sheet_ranges['A5'].value
    localData = local.split(':')

    origin = ''
    destination = ''

    if "ORG" in localData[0]:
        origin = localData[1].strip()
    else:
        destination = localData[1].strip()


    max_col = sheet_ranges.rows.gi_frame.f_locals['max_col']
    max_row = sheet_ranges.rows.gi_frame.f_locals['max_row']

    for row in sheet_ranges.iter_rows(min_row=9, max_col=max_col, max_row=max_row):
        fields = []
        for cell in row:
            fields.append(cell) 
            
        if origin:
            fields.insert(1,origin)
        else:
            fields.insert(2,destination)
        service = Service(fields)
        if service.is_valid_service():
            result = [serviceToFilter for serviceToFilter in services if service.paxName == serviceToFilter.paxName]
            if result:
                indx = services.index(result[0])
                service_with_connection = put_in_connection_of_service(service, services[indx])
                services.remove(result[0])
                services.append(service_with_connection)
            else:
                services.append(service)
                
    services_to_excel.convert_to_excel(services)


def is_valid_sheet_name(sheetName):
    keys = ['SALIDAS', 'LLEGADAS']
    for key in keys:
        if key in sheetName:
            return True
    return False
            



def put_in_connection_of_service(service, connection):
    service.flightConnectionNumber = connection.flightNumber
    service.endZone = 'SALAS'
    return service


if __name__ == '__main__':
    main()
