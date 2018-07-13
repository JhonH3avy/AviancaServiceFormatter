from openpyxl import Workbook

def convert_to_excel(services = []):
    wb = Workbook()
    ws = wb.active
    ws.title = "Servicios Avianca"
    ws['A1'] = 'flightId'
    ws['B1'] = 'airline'
    ws['C1'] = 'originCity'
    ws['D1'] = 'destinyCity'
    ws['E1'] = 'startDate'
    ws['F1'] = 'flightConnectionId'
    ws['G1'] = 'paxName'
    ws['H1'] = 'paxReservationNumber'
    ws['I1'] = 'startZone'
    ws['J1'] = 'endZone'
    ws['K1'] = 'connectionGate'
    ws['L1'] = 'serviceType'
    ws['M1'] = 'reserved'
    ws['N1'] = 'authorizedBy'

    total_services = len(services)
    for i in range(2,total_services):
        ws['A' + str(i)] = services[i].flightNumber
        ws['B' + str(i)] = services[i].airline
        ws['C' + str(i)] = services[i].origin
        ws['D' + str(i)] = services[i].destination
        ws['E' + str(i)] = services[i].startDate
        ws['F' + str(i)] = services[i].flightConnectionNumber
        ws['G' + str(i)] = services[i].paxName
        ws['H' + str(i)] = services[i].paxReservationNumber
        ws['I' + str(i)] = services[i].startZone
        ws['J' + str(i)] = services[i].endZone
        ws['K' + str(i)] = services[i].connectionGate
        ws['L' + str(i)] = services[i].serviceType
        ws['M' + str(i)] = services[i].reserved
        ws['N' + str(i)] = services[i].authorizedBy


    return wb