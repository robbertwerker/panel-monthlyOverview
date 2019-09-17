import os, csv, datetime

def main():
    csvFiles = getCSVFile()

    crossPanelData = []
    order = []
    for i in csvFiles:
        panelistData = extractRelevantData(i)
        subscriptionDetails = calculateSubscription(panelistData)
        activityDetails = calculateActivity(panelistData)
        combinedDict = {}
        combinedDict.update(activityDetails)
        combinedDict.update(subscriptionDetails)
        crossPanelData.append(combinedDict)
        order.append(i[13:])
        # print(order)
    return [crossPanelData, order]

#function to find a csv file in the current working directory. If more than one .csv file is present, it will use the first one
def getCSVFile():
    fileList = os.listdir()
    filteredList = filter(lambda x: x.endswith('.csv'), fileList)
    filteredList = list(filteredList)
    csvFiles = filteredList
    return csvFiles

#function to clean up the csv file (removing newline), get the indexes of the headers that we need
def extractRelevantData(csvFile):
    with open(csvFile, 'r', newline='') as dataFile:
        dataFile_reader = csv.reader(dataFile, delimiter=',', quoting=csv.QUOTE_MINIMAL, lineterminator='\n')

        #clean up headers
        headers = []
        relevantData = []
        for i, line in enumerate(dataFile_reader):
            if i == 0:
                for cell in line:
                    headers.append(cell.replace("\n", ""))
                subscribedIndex = headers.index('Subscribed?')
                emailokIndex = headers.index('Email valid?')
                lastParticipationIndex = headers.index('Last survey start time')
                dateUnsubscribedIndex = headers.index('Time unsubscribed')
                joinedIndex = headers.index('Join date')
                # print(subscribedIndex, emailokIndex, lastParticipationIndex, dateUnsubscribedIndex, joinedIndex)
        
        #find indexes of relevant metrics and add to relevantData array
            else:
                panelistData = {
                    "id": line[0],
                    "subscribed": line[subscribedIndex],
                    "emailok": line[emailokIndex],
                    "lastParticipation": line[lastParticipationIndex],
                    "dateUnsubscribed": line[dateUnsubscribedIndex],
                    "joined": line[joinedIndex]
                }
                relevantData.append(panelistData)
        return relevantData

#function to turn datetime object to a string
def dateObject(str, extend):
    if extend:
        return datetime.datetime.strptime(str, "%Y-%m-%d %H:%M:%S")
    else:
        return datetime.datetime.strptime(str, "%Y-%m-%d")

#function to get the panel subscription details
def calculateSubscription(panelistData):
    today = datetime.datetime.now()
    IngeschrevenUitgeschrevenAfgelopenMaand = []
    actualDaysUnsubscribed = 0
    daysSinceJoining = 0
    IngeschrevenAfgelopenMaand = 0
    IngeschrevenAfgelopenKwartaal = 0
    IngeschrevenAfgelopenHalfJaar = 0
    IngeschrevenAfgelopenJaar = 0
    UitgeschrevenAfgelopenMaand = 0
    for row in panelistData:
        #create days unsubscribed variable
        date_unsubscribed = row["dateUnsubscribed"]
        date_joined = row["joined"]
        subscribed = row["subscribed"]
        if len(date_unsubscribed) > 0:
            dateUnsubscribedObject = dateObject(date_unsubscribed, True)
            daysUnsubscribed = today - dateUnsubscribedObject
            if ' ' in str(daysUnsubscribed):
                actualDaysUnsubscribed = int(str(daysUnsubscribed).split(" ")[0])
            else:
                actualDaysUnsubscribed = 0
        else:
            actualDaysUnsubscribed = 0
        #calculate days since joining
        if len(date_joined) > 0:
            dateJoinedObject = dateObject(date_joined, False)
            dateSinceJoining = today - dateJoinedObject
            if ' ' in str(dateSinceJoining):
                daysSinceJoining = int(str(dateSinceJoining).split(" ")[0])
            else:
                daysSinceJoining = 0
            

        if subscribed == 'no' and daysSinceJoining < 30 and actualDaysUnsubscribed < 30:
            IngeschrevenUitgeschrevenAfgelopenMaand.append(row)
        if daysSinceJoining < 30:
            IngeschrevenAfgelopenMaand += 1
        if daysSinceJoining < 90:
            IngeschrevenAfgelopenKwartaal += 1
        if daysSinceJoining < 180:
            IngeschrevenAfgelopenHalfJaar += 1
        if daysSinceJoining < 365:
            IngeschrevenAfgelopenJaar += 1
        if actualDaysUnsubscribed < 30 and row["subscribed"] != "yes" and actualDaysUnsubscribed != 0:
            UitgeschrevenAfgelopenMaand += 1

    subscriptionDetails = {
        "IngeschrevenAfgelopenJaar": IngeschrevenAfgelopenJaar,
        "IngeschrevenAfgelopenHalfJaar": IngeschrevenAfgelopenHalfJaar,
        "IngeschrevenAfgelopenKwartaal": IngeschrevenAfgelopenKwartaal,
        "IngeschrevenAfgelopenMaand": IngeschrevenAfgelopenMaand,
        "UitgeschrevenAfgelopenMaand": UitgeschrevenAfgelopenMaand,
        "IngeschrevenUitgeschrevenAfgelopenMaand": len(IngeschrevenUitgeschrevenAfgelopenMaand),
    }
    return subscriptionDetails

#function used to check the panel activity over the last period
def calculateActivity(panelistData):
    today = datetime.datetime.now()
    daysSinceLastParticipation = 0
    ACTIEFDEELmaand = 0
    ACTIEFDEELkwartaal = 0
    ACTIEFDEELjaar = 0
    for row in panelistData:
        last_participation = row["lastParticipation"]
        if len(last_participation) > 0:
            lastParticipationObject = dateObject(last_participation, True)
            daysSinceLastParticipation = today - lastParticipationObject

            #if the time since last_participation is less than 1 day, the format is different. The line below is used to circumvent this issue.
            if ' ' in str(daysSinceLastParticipation):
                daysSinceLastParticipation = str(daysSinceLastParticipation).split(" ")[0]
            else:
                daysSinceLastParticipation = 0
            #fill variables
            if int(daysSinceLastParticipation) < 30:
                ACTIEFDEELmaand += 1
            if int(daysSinceLastParticipation) < 90:
                ACTIEFDEELkwartaal += 1
            if int(daysSinceLastParticipation) < 365:
                ACTIEFDEELjaar += 1
    activityDetails = {
        "ACTIEFDeelmaand": ACTIEFDEELmaand,
        "ACTIEFDEELkwartaal": ACTIEFDEELkwartaal,
        "ACTIEFDEELjaar": ACTIEFDEELjaar
    }
    return activityDetails

if __name__ == "__main__":
    main()