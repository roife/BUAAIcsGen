import xlrd
import re
import datetime


data = xlrd.open_workbook('schedule.xls', 'rb')
ics = open('schedule.ics', 'w', encoding='utf-8')
table = data.sheets()[0]

termBeginDate = datetime.date(2019, 9, 2)


def eventSetGen(table):
    eventSet = []
    beginTime = {1: '0800', 2: '0850', 3: '0950', 4: '1040', 5: '1130',
                 6: '1400', 7: '1450', 8: '1550', 9: '1640', 10:'1730',
                 11:'1900', 12:'1950', 13:'2040', 14:'2130'}
    endTime = {1: '0845', 2: '0935', 3: '1035', 4: '1125', 5: '1215',
               6: '1445', 7: '1535', 8: '1635', 9: '1725', 10:'1815',
               11:'1945', 12:'2035', 13:'2125', 14:'2215'}
    for i in range(2, 8): # column
        for j in range(2, 8): # row
            if(table.row_values(j)[i] != ''):
                cell = table.row_values(j)[i]
                print(cell,'\n')
                # classNum = len(re.findall(r'第(([0-9]+)|，)+节', cell))
                classSet = re.split(r'节</br>', cell)
                for classi in classSet:
                    if(classi[-1] != '节'):
                        classi.join('节')
                    blockNum = len(re.findall(r'\[[0-9]*-[0-9]*]|\[[0-9]*]', classi))
                    eventWeekday = i-1
                    eventTimeSet = list(map(int, (re.findall(r'([0-9]+)[节|，]', classi))))
                    if(blockNum == 1):
                        eventName = re.match(r'.*</br>', classi).group()[:-5]
                        eventBeginWeek = re.search(r'\[[0-9]*', classi).group()[1:]
                        # [WEEK] and [BEGIN-END]
                        if(re.search(r'-.*]', classi) is None):
                            eventEndWeek = re.search(r'\[.*]', classi).group()[1:-1]
                        else:
                            eventEndWeek = re.search(r'-.*]', classi).group()[1:-1]
                        eventPlace = re.search(r'周.*\n', classi).group()[1:-1]
                        # print(classi, '\n', eventTimeSet, '\n')
                        eventDescription = re.sub("(</br>)|\\n", " ", classi)
                        event = [eventName, eventBeginWeek, eventEndWeek,
                                 beginTime[eventTimeSet[0]], endTime[eventTimeSet[-1]], eventPlace, eventWeekday, eventDescription]
                        eventSet.append(event)
                    else:
                        subEventWeek = (re.findall(r'\[[0-9]*-[0-9]*\]|\[[0-9]*]', classi))
                        subEventPlace = (re.findall(r'周.*\n', classi))
                        for k in range(blockNum):
                            subEventName = classi[:classi.find('</br>')]
                            subEventBeginWeek = re.match(r'\[[0-9]*', subEventWeek[k]).group()[1:]
                            # [WEEK] and [BEGIN-END]
                            if(re.search(r'-.*]', subEventWeek[k]) is None):
                                subEventEndWeek = re.search(r'\[.*]', subEventWeek[k]).group()[1:-1]
                            else:
                                subEventEndWeek = re.search(r'-.*]', subEventWeek[k]).group()[1:-1]
                            eventPlace = re.search(r'周.*\n', subEventPlace[k]).group()[1:-1]
                            eventDescription = re.sub("(</br>)|\\n", " ", classi)
                            subEvent = [subEventName, subEventBeginWeek, subEventEndWeek,
                                        beginTime[eventTimeSet[0]], endTime[eventTimeSet[-1]], eventPlace, eventWeekday, eventDescription]
                            eventSet.append(subEvent)
    return eventSet


eventSet = eventSetGen(table)
print(eventSet)


def calcBeginDate(beginDate, intervalWeek, weekday):
    interval = datetime.timedelta(
        weeks=int(intervalWeek-1))+datetime.timedelta(days=weekday-1)
    ansDate = ''.join(re.split('-', str(beginDate+interval)))
    return ansDate


def payloadGen(eventSet):
    eventBeginDateSet = list(
        map(lambda x: calcBeginDate(termBeginDate, int(x[1]), x[6]), eventSet))

    eventLastWeek = list(map(lambda x: int(x[2])-int(x[1])+1, eventSet))

    event = [{
        'DTSTART': eventBeginDateSet[i]+'T'+eventSet[i][3]+'00',
        'DTEND':eventBeginDateSet[i]+'T'+eventSet[i][4]+'00',
        'SUMMARY':eventSet[i][0],
    } for i in range(len(eventSet))]
    # print(event[0]['DTSTART'], event[0]['DTEND'])

    payload = 'BEGIN:VCALENDAR\nVERSION:2.0\nCALSCALE:GREGORIAN\nBEGIN:VTIMEZONE\nTZID:Asia/Shanghai\nTZURL:http://tzurl.org/zoneinfo-outlook/Asia/Shanghai\nX-LIC-LOCATION:Asia/Shanghai\nBEGIN:STANDARD\nTZOFFSETFROM:+0800\nTZOFFSETTO:+0800\nTZNAME:CST\nDTSTART:19700101T000000\nEND:STANDARD\nEND:VTIMEZONE'

    for i in range(len(event)):
        vEvent = '''\nBEGIN:VEVENT\nDTSTAMP:20190902T123110Z\nDTSTART;TZID=Asia/Shanghai:''' + \
            event[i]['DTSTART']+'\nDTEND;TZID=Asia/Shanghai:'+event[i]['DTEND']+'\nRRULE:FREQ=WEEKLY;COUNT=' + \
            str(eventLastWeek[i])+'\nSUMMARY:'+eventSet[i][0] + \
            '\nLOCATION:'+eventSet[i][5]+'\nDESCRIPTION:'+eventSet[i][7] + \
            '\nBEGIN:VALARM'+'\nTRIGGER:-PT15M'+'\nREPEAT:1'+'\nDURATION:PT1M'+'\nEND:VALARM'+'\nEND:VEVENT'
        payload += vEvent
    payload += '\nEND:VCALENDAR'
    # print(payload)
    return payload

ics.write(payloadGen(eventSet))
ics.close()