import arrow

class Service(object):
    """Flight object to be inserted in database"""
    def __init__(self, fields = []):
        self.airline = 'AVA'
        if isinstance(fields[0].value, str) and  isinstance(fields[3].value, str):
            self.flightNumber = fields[0].value + fields[3].value
        else:
            self.flightNumber = None
        if isinstance(fields[1], str):
            self.origin = fields[1]
            self.destination = fields[2].value
        else:
            self.origin = fields[1].value
            self.destination = fields[2]
        if self.origin == 'BOG':
            if fields[4].value:
                self.startDate = arrow.get(fields[4].value, 'America/Bogota').to('utc').format('YYYY-MM-DD HH:mm') 
            else:
                self.startDate = None
            self.startZone = 'MODULOS'
            self.endZone = 'SALAS'
        else:
            if fields[5].value:
                self.startDate = arrow.get(fields[5].value, 'America/Bogota').to('utc').format('YYYY-MM-DD HH:mm')            
            else:
                self.startDate = None
            self.startZone = 'GATE'
            self.endZone = 'SALIDA'
        self.paxName = fields[6].value
        self.serviceType = None
        if fields[8].value:
            self.serviceType = 'WCOB'
        elif fields[9].value:
            self.serviceType = 'WCMP'
        elif fields[10].value:
            self.serviceType = 'WCHS'
        elif fields[12].value:
            self.serviceType = 'WCHC'
        elif fields[11].value:
            self.serviceType = 'WCHR'
        try:
            if self.serviceType is None and fields[13].value:
                self.serviceType = 'WCBW'
        except IndexError:
            print('No WCBW Column for this service')
        self.flightConnectionNumber = ''
        self.paxReservationNumber = ''
        self.connectionGate = ''
        self.reserved = 1
        self.authorizedBy = ''


    def is_valid_service(self):
        if not self.flightNumber:
            return False
        if not self.startDate:
            return False
        if not self.paxName:
            return False
        if not self.serviceType:
            return False
        if not self.origin:
            return False
        if not self.destination:
            return False
        return True



