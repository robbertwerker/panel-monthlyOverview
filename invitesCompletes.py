import requests
import json
import pprint
import csv
pp = pprint.PrettyPrinter(indent=4)
baseURL = "https://panel.noviopinion.com/api.pro?method="

def main(panel, month, year):
    print("Welcome to this Kinesis data fetching tool. It'll help you retrieve monthly data from Kinesis quickly. Please follow the instructions on the screen.\n")
    getSesKeyObject = getSesKey(panel)
    sesKey = getSesKeyObject["sesKey"]
    panelCode = getSesKeyObject["panelCode"]
    projectList = getProjectList(sesKey)
    relevantProjectsObject = filterProjects(sesKey, projectList, month, year)
    year = relevantProjectsObject["year"]
    monthTwoDigits = relevantProjectsObject["monthTwoDigits"]
    relevantProjects = relevantProjectsObject["data"]
    projectData = getProjectData(sesKey, relevantProjects)
    # print(projectData)
    endData = writeToCSV(projectData, panelCode, year, monthTwoDigits)
    return endData
    # print(input("Success! Writing to csv file complete. You can find the file you've just created in the same folder as the .exe file. \nThis is the end of the script. Press enter to close the terminal."))

# get a session key
def getSesKey(panel):
    # while True:
    #     panelCode = panel 
    #     if panelCode not in ["2", "39", "49", "45", "54", "31"]:
    #         print('Please enter a valid panel code')
    #     else:
    #         break
    data = json.dumps({
        "username": "apiUser",
        "password": "Welkom01!",
        "panelid": panel
    })
    method = "integration.auth.login&"
    dataURL = "data={data}".format(data=data)
    sesKey = requests.post(baseURL + method + dataURL)
    sesKey = json.loads(sesKey.content.decode('utf-8'))["data"]['sesKey']
    return {"sesKey": sesKey, "panelCode":panel}

# get list of all projects in that portal
def getProjectList(sesKey):
    data = json.dumps({
        'sesKey': sesKey
    })
    method = "integration.project.listing&"
    dataURL = "data={data}".format(data=data)
    rawProjectList = requests.post(baseURL + method + dataURL)
    projectListData = json.loads(
        rawProjectList.content.decode('utf-8'))["data"]
    return projectListData

# filter project to only include those that were created in the relevant month
def filterProjects(sesKey, projectList, month, year):

    # # check if user input is correct
    # while True:
    #     if month not in ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]:
    #         print('Please enter a number between 1 and 12 (there are 12 months in a year)')
    #     else:
    #         break
    # monthTwoDigits = '0' + month if len(month) == 1 else month
    filteredArray = []
    for project in projectList:
        if project['created'] >= '{year}-{month}-01 00:00:00'.format(month=month, year=year) and project["created"] <= '{year}-{month}-31 23:59:59'.format(month=month, year=year):
            filteredArray.append(project)
    print(len(filteredArray), 'projects found. Please wait for the data to be retrieved...')
    return {"data": filteredArray, "monthTwoDigits": month, "year":year}

#get data of the filtered projects
def getProjectData(sesKey, relevantProjects):
    allProjectData =[]
    for i, project in enumerate(relevantProjects):
        # select project
        data = json.dumps({
            'sesKey': sesKey,
            'projectid': project["id"]
        })
        method = "integration.project.select&"
        dataURL = "data={data}".format(data=data)
        requests.post(baseURL + method + dataURL)
        print("Projects loading, please wait: ", len(relevantProjects) - i)
        # read project data
        data = json.dumps({
            'sesKey': sesKey,
            'type': 'statistics'
        })
        method = "integration.project.read&"
        dataURL = "data={data}".format(data=data)
        rawProjectData = requests.post(baseURL + method + dataURL)
        projectData = json.loads(
            rawProjectData.content.decode('utf-8'))["data"]
        # pp.pprint(projectData)

        kinesisProjectData = {
            "projectid": project['id'],
            "projectName": project["name"],
            "invited": projectData["completion"]["total"],
            "started": projectData["response"]["responded"]["count"],
            "startPerc": projectData["response"]["responded"]["percent"],
            "completed": projectData["completion"]["completed"]["count"],
            "complPerc": projectData["completion"]["completed"]["percent"],
            "screened": projectData["completion"]["profile"]["count"],
            "dropout": projectData["completion"]["started"]["count"]
        }
        allProjectData.append(kinesisProjectData)

    return allProjectData
   
# write results to csv file
def writeToCSV(projectData, panelCode, year, monthTwoDigits):
    panelConvert ={
        "2": "Pollland",
        "39": "Orange Buddies NL",
        "49": "Orange Buddies BE",
        "45": "Orange Buddies UK",
        "54": "EuroQuestions",
        "31": "Clubdoel"
    }

    panelName = panelConvert[panelCode]
    headers = ["projectid", "projectName","invited", "started", "started%", "completed", "completed%", "screened", "dropout"]
    # file = "{panelName} data {year}{month}".format(panelName=panelName, year=year, month=monthTwoDigits)
    # if not file.endswith('.csv'):
    #     file = file + '.csv'
    # with open(file, 'w', newline='') as projectFile:
    #     projectFile_writer = csv.writer(projectFile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator='\n')
    #     projectFile_writer.writerow(headers)
    invitedCount = 0
    startedCount = 0
    completedCount = 0
    screenedCount = 0
    dropoutCount = 0
    for i, project in enumerate(projectData):
        # projectFile_writer.writerow([project["projectid"],project["projectName"],project["invited"],project["started"],project["startPerc"],project["completed"],project["complPerc"], project["screened"],project["dropout"]])
        invitedCount += int(project["invited"])
        startedCount += int(project["started"])
        completedCount += int(project["completed"])
        screenedCount += int(project["screened"])
        dropoutCount += int(project["dropout"])
    panelScores = {
        "invited": invitedCount,
        "started": startedCount,
        "completed": completedCount,
        "screened": screenedCount,
        "dropout": dropoutCount
    }
    return panelScores
    # projectFile_writer.writerow(['', 'TOTAL', invitedCount, startedCount, '', completedCount, '', screenedCount, dropoutCount])
    # projectFile.close()


