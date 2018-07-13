"""
A sample for working with the Pyvot library for Excel.
"""

from os import path
from openpyxl import load_workbook
from Service import Service
import services_to_excel
import sys

OK = 0
FILE_LOADING_ERROR = 1
FILE_FORMAT_ERROR = 2
FILE_ACCESS_PERMISSION_ERROR = 3

def main():
    """Loads and modifies a sample spreadsheet in Excel."""
    flags = {}
    flag = None
    for arg in sys.argv:
        if '-' in arg:
            flag = arg
        else:
            if flag:
                flags[flag] = arg
                flag = None

    
    if flags['-p']:
        workbook_path = flags['-p']
        print('Using custom path')
    else:
        workbook_path = path.join(path.dirname(__file__), 'service_data.xlsx')
    try:
        workbook = load_workbook(workbook_path, read_only=True)
        print('The file has been loaded successfully')
    except FileNotFoundError:
        print('The file has not been found in ' + workbook_path);
        sys.exit(FILE_LOADING_ERROR)
    except FileExistsError:
        print('The file ' + workbook_path + ' doesn\'t exist')
        sys.exit(FILE_LOADING_ERROR)


    sheet_names = workbook.sheetnames
    servicesToParse = None
    for sheet in sheet_names:
        print('Getting services of sheet ' + sheet)
        servicesToParse = process_sheet(sheet, workbook, servicesToParse)

    print(str(len(servicesToParse)) + ' services has been created from file ' + workbook_path)
    normalized_excel = services_to_excel.convert_to_excel(servicesToParse)

    resultFileName = 'services.xlsx'
    try:
        normalized_excel.save(resultFileName)
        print('Excel file created successfully')
    except PermissionError:
        print('Program do not have permission to access ' + path.dirname(__file__) + resultFileName)
        sys.exit(FILE_ACCESS_PERMISSION_ERROR)
    print('Job Done...')
    sys.exit(OK)    



def process_sheet(sheet, workbook, services = []):   
    if services is None:
        services = []
    if not is_valid_sheet_name(sheet):
        print(sheet + ' is not a valid sheet name.')
        return services

    sheet_ranges = workbook[sheet]    

    try:
        assert sheet_ranges['A5'].value is not None
        local = sheet_ranges['A5'].value
        localData = local.split(':')
    except AssertionError:
        print('File is not in correct format. Cell "A5" has no destination nor origin information')
        sys.exit(FILE_FORMAT_ERROR)

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
        
    print('Services for ' + sheet + ' has been processed')   
    return services
    


def is_valid_sheet_name(sheetName):
    keys = ['SALIDAS', 'SALIENDO', 'LLEGADAS', 'LLEGANDO']
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
