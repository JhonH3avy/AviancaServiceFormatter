import arrow

class Service(object):
    """Flight object to be inserted in database"""
    def __init__(self, fields = []):
        self.airline = 'AVA'
        if isinstance(fields['carrier'], str) and  isinstance(fields['flightNumber'], str):
            self.flightNumber = fields['carrier'] + fields['flightNumber']
        else:
            self.flightNumber = None
        self.origin = fields['origin']
        self.destination = fields['destination']
        if self.origin == 'BOG':
            if fields['departure']:
                self.startDate = arrow.get(fields['departure'], 'America/Bogota').to('utc').format('YYYY-MM-DD HH:mm') 
            else:
                self.startDate = None
            self.startZone = 'MODULOS'
            self.endZone = 'SALAS'
        else:
            if fields['arrival']:
                self.startDate = arrow.get(fields['arrival'], 'America/Bogota').to('utc').format('YYYY-MM-DD HH:mm')            
            else:
                self.startDate = None
            self.startZone = 'GATE'
            self.endZone = 'SALIDA'
        self.paxName = fields['paxName']
        self.serviceType = None
        if 'WCOB' in fields.keys() and fields['WCOB']:
            self.serviceType = 'WCOB'
        elif 'WCMP' in fields.keys() and fields['WCMP']:
            self.serviceType = 'WCMP'
        elif 'WCHS' in fields.keys() and fields['WCHS']:
            self.serviceType = 'WCHS'
        elif 'WCHC' in fields.keys() and fields['WCHC']:
            self.serviceType = 'WCHC'
        elif 'WCHR' in fields.keys() and fields['WCHR']:
            self.serviceType = 'WCHR'
        elif 'WCBW' in fields.keys() and fields['WCBW']:
            self.serviceType = 'WCBW'
        elif 'WEAP' in fields.keys() and fields['WEAP']:
            self.serviceType = 'WEAP'
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


    def is_arriving(self):
        return self.destination == 'BOG'



