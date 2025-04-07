import requests
import sys
from bs4 import BeautifulSoup
from itertools import combinations
#import math
import pandas as pd
import os
import time

class Course:
    """Stores all relevant properties of a course, including the course number, the course name,
    Professor, days of lecture, meeting times, and section"""

    def __init__(self, course_num, course_name, prof, days, time, niceTime, section, CRN, seats, waitlist):
        self.course_num = course_num
        self.course_name = course_name
        self.prof = prof
        """Sets days of course"""
        if len(days) > 1:
            self.days = days[0]
            self.labDays = days[1]
        else:
            self.days = days
            self.labDays = -1

        """Sets times of course and lab time if applicable"""
        self.timeBegin = time[0][0]
        self.timeEnd = time[0][1]
        if len(time) > 1:

            self.labTimeBegin = time[1][0]
            self.labTimeEnd = time[1][1]
            self.niceTime = niceTime[0]
            self.niceLabTime = niceTime[1]
        else:
            if time[0][0] == -1:
                self.niceTime = -1
            self.labTimeBegin = -1
            self.labTimeEnd = -1
            self.niceLabTime = "NA"
            self.niceTime = niceTime


        self.elective = False
        if course_num[len(course_num) - 3] == 'V':
            self.hours = "Research"
        else:
            if section.isalpha():
                self.hours = 0
            else:
                self.hours = int(course_num[len(course_num) - 3])
        self.section = section

        if self.section.isalpha():
            self.isLab = True
        elif not self.section.isalpha():
            self.isLab = False
        else:
            self.isLab = False
        self.CRN = CRN
        self.seats = seats
        self.waitlist = waitlist

    def printCourseInfo(self):
        """Prints relevant info stored in course class"""
        print()
        print(self.course_num)
        print("Course name: " + self.course_name)
        print("Credit hours: " + str(self.hours))
        print("Professor: " + self.prof)
        if self.labDays != -1:
            print("Class days: " + self.days)
            print("Lab days: " + str(self.labDays))
        else:
            print("Class days: " + self.days[0])


        if self.labTimeBegin != -1:
            print("Class time: " + self.niceTime)
            print("Lab time: " + self.niceLabTime)
        else:
            if self.niceTime != "Online":
                print("Class time: " + self.niceTime[0])

        if self.elective:
            print("Elective")
        else:
            print("Required Class")
        print("Section: " + self.section)
        print("CRN: " + self.CRN)
        print("Seats Available: " + str(self.seats))
        print("Waitlist: " + str(self.waitlist))
        print()

    def writeInfo(self,file):
        """Writes relevant info stored in course class"""

        file.write("\n" + self.course_num + "\n")
        file.write("Course name: " + self.course_name + "\n")
        file.write("Credit hours: " + str(self.hours) + "\n")
        file.write("Professor: " + self.prof + "\n")
        if self.labDays != -1:
            file.write("Class days: " + self.days + "\n")
            file.write("Lab days: " + str(self.labDays) + "\n")
        else:
            file.write("Class days: " + self.days[0] + "\n")

        if self.labTimeBegin != -1:
            file.write("Class time: " + self.niceTime + "\n")
            file.write("Lab time: " + self.niceLabTime + "\n")
        else:
            if self.niceTime != "Online":
                file.write("Class time: " + self.niceTime[0] + "\n")

        if self.elective:
            file.write("Elective" + "\n")
        else:
            file.write("Required Class" + "\n")
        file.write("Section: " + self.section + "\n")
        file.write("CRN: " + self.CRN + "\n")
        file.write("Seats Available: " + str(self.seats) + "\n")
        file.write("Waitlist: " + str(self.waitlist) + "\n")


    def getData(self):
        """Retrieves relevant info from course class and stores in dictionary.
        Used primarily for exporting course data to an Excel file"""
        data = {
            "Course Number": self.course_num,
            "Course Name": self.course_name,
            "Section": self.section,
            "CRN": self.CRN,
            "Professor": self.prof,
            "Seats Available": self.seats,
            "Waitlist": self.waitlist,
        }
        if self.labDays != -1:
            data["Class Days"] = self.days
            data["Class Times"] = self.niceTime
            data["Lab Days"] = self.labDays
            data["Lab Times"] = self.niceLabTime
        else:
            data["Class Days"] = self.days[0]
            data["Class Times"] = self.niceTime[0]
        return data



class Schedules:
    """Class that stores an array of courses"""
    def __init__(self):
        self.schedule_list = []

    def printSchedule(self):
        """Prints info of courses within every schedule"""
        for m in self.schedule_list:
            print("==============================================")
            for n in m:
                n.printCourseInfo()
            print("==============================================")

    def writeToFile(self, file):
        """Writes relevant info stored in course class to file"""
        file.write(str(len(self.schedule_list)) + " valid schedules\n")
        p = 1
        for m in self.schedule_list:
            file.write("==============================================\n")
            file.write("Schedule " + str(p) + "\n")
            for n in m:
                n.writeInfo(file)
            file.write("==============================================")
            p = p + 1
    def getNumSchedules(self):
        """Returns number of schedules stored in class"""
        return len(self.schedule_list)

    def addSchedule(self,schedule):
        """Appends a schedule to schedule list"""
        self.schedule_list.append(schedule)

    def getData(self):
        """Calls to getData function for each course within each schedule and retrieves
        relevant data. Used to export schedule data to an Excel file"""
        data = []
        for m in self.schedule_list:
            scheduleData = []
            for n in m:
                scheduleData.append(n.getData())
            data.append(scheduleData)
        return data




def convertToMilitaryTime(time):
    """Converts AM and PM times to 24hr format. Useful when checking if course times
    overlap when going through schedule combinations"""
    times = []
    for m in time:
        lower = m[0:6]
        upper = m[9:]

        if lower[4:6] == 'PM' and lower[0] == '0':
            lower = int(lower[0:4]) + 1200
            if lower > 2400:
                lower = lower - 2400
        else:
            lower = int(lower[0:4])
        if upper[4:6] == 'PM' and upper[0] == '0':
            upper = int(upper[0:4]) + 1200
            if upper > 2400:
                upper = upper - 2400
        else:
            upper = int(upper[0:4])
        times.append([lower, upper])
    return times

def makeHeader(i):
    """Creates header for Excel file"""
    header = {
        "Course Number": "Schedule " + str(i),
        "Course Name": "",
        "Section": "",
        "CRN": "",
        "Professor": "",
        "Seats Available": "",
        "Waitlist": "",
        "Class Days": "",
        "Class Times": "",
        "Lab Days": "",
        "Lab Times": ""

    }
    return header

def letterToNum(letter):
    return ord(letter[0]) - 64


stime = time.time()
#############################################################################################################
noInputFlag = 0
blankFlag = 0

"""Input"""
inputFile = open("UserInput.txt")

inputFileContents = inputFile.read()
inputFileContents = inputFileContents.split('\n')

inputFile.close()

temp = inputFileContents[0]

"""Check for invalid inputs(enter or "")"""
tempReqClasses = temp[temp.find("[")+1:temp.find("]")]
if "enter" in tempReqClasses:
    noInputFlag += 1
    print("Error: Please replace \"enter\" in Required Classes with a valid input")

if tempReqClasses == "":
    blankFlag += 1
    print("Error: Required Classes left blank, please input valid classes")

temp = inputFileContents[2]
tempElectiveClasses = temp[temp.find("[")+1:temp.find("]")]
if "enter" in tempElectiveClasses:
    noInputFlag += 1
    print("Error: Please replace \"enter\" in Elective Classes with a valid input")
    print("If you do not want to input Elective Classes, delete \"enter\" and leave field blank")

temp = inputFileContents[4]
maxHours = temp[temp.find("[")+1:temp.find("]")]
if "enter" in maxHours:
    noInputFlag += 1
    print("Error: Please replace \"enter\" in Max Hours with a valid input")
else:
    maxHours = int(maxHours)

if maxHours == "":
    blankFlag += 1
    print("Error: Max Hours field left blank, please input a valid number of hours")

temp = inputFileContents[6]
term = temp[temp.find("[")+1:temp.find("]")]
if "enter" in term:
    noInputFlag += 1
    print("Error: Please replace \"enter\" in Term with a valid input")
if term == "":
    blankFlag += 1
    print("Error: Term left blank, please input a valid term")

temp = inputFileContents[8]
year = temp[temp.find("[")+1:temp.find("]")]
if "enter" in year:
    noInputFlag += 1
    print("Error: Please replace \"enter\" in Year with a valid input")
if year == "":
    blankFlag += 1
    print("Error: Year left blank, please input a valid year")

temp = inputFileContents[10]
saveToExcel = temp[temp.find("[")+1:temp.find("]")]
if "enter" in saveToExcel:
    noInputFlag += 1
    print("Error: Please replace \"enter\" in Save to Excel with a valid input")
if saveToExcel == "":
    blankFlag += 1
    print("Error: Export to Excel left blank, please input Yes or No")
if "yes" in saveToExcel.lower():
    saveToExcel = True
else:
    saveToExcel = False

temp = inputFileContents[12]
saveToTxt = temp[temp.find("[")+1:temp.find("]")]
if "enter" in saveToTxt:
    noInputFlag += 1
    print("Error: Please replace \"enter\" in Save to txt with a valid input")
if saveToTxt == "":
    blankFlag += 1
    print("Error: Export to Excel left blank, please input Yes or No")
if "yes" in saveToTxt.lower():
    saveToTxt = True
else:
    saveToTxt = False


if noInputFlag + blankFlag > 0:
    print()
    print("Please read UserInputInstructions.txt then edit UserInput.txt to run program")
    time.sleep((noInputFlag + blankFlag) * 5)
    sys.exit()

#############################################################################################################
"""Store names of course in arrays"""
reqClassesArr = tempReqClasses.split(", ")
electiveClassesArr = tempElectiveClasses.split(", ")

reqPreferredProf = []
reqPreferredProfClass = []

electivePreferredProf = []
electivePreferredProfClass = []

"""Searches for professor names enclosed in parenthesis.
Stores names of professors(if found) in a array and removes the parenthesis
and professor name from class name in class array"""
for i in range(len(reqClassesArr)):
    temp = reqClassesArr[i].find("(")
    if temp != -1:
        reqPreferredProf.append(reqClassesArr[i][temp+1:len(reqClassesArr[i])-1])
        reqClassesArr[i] = reqClassesArr[i][0:temp]
        reqPreferredProfClass.append(reqClassesArr[i])

for i in range(len(electiveClassesArr)):
    temp = electiveClassesArr[i].find("(")
    if temp != -1:
        electivePreferredProf.append(electiveClassesArr[i][temp+1:len(electiveClassesArr[i])-1])
        electiveClassesArr[i] = electiveClassesArr[i][0:temp]
        electivePreferredProfClass.append(electiveClassesArr[i])


preferredProf = reqPreferredProf + electivePreferredProf
preferredProfClass = reqPreferredProfClass + electivePreferredProfClass

if electiveClassesArr[0] != "":
    allClassesArr = reqClassesArr + electiveClassesArr
else:
    allClassesArr = reqClassesArr

#############################################################################################################

if term.lower() == "spring":
    term = year + "10"
elif term.lower() == "fall":
    term = year + "30"
else:
    print("Error: Term must be spring or fall")
    time.sleep(5)
    sys.exit()

classes = []
noCourseNumArr = []
reqSectionClasses = []

"""Web scrapes data related to each input classes"""
print("Collecting data...")
for x in allClassesArr:
    prefix = x[0:x.find(" ")]
    num = x[x.find(" ")+1:]
    url = "https://www1.baylor.edu/scheduleofclasses/Results.aspx?Term="
    url = url + term
    url = url + "&College=Z&Prefix="
    url = url + prefix
    url = url + "&StartCN="
    url = url + num
    url = url + "&EndCN="
    url = url + num
    url = url + "&Status=Z&Days=Z&Instructor=&IsMini=false&OnlineOnly=0&POTerm=Z&CourseAttr=Z&Sort=SN"

    print(url)

    page = requests.get(url)
    soup = BeautifulSoup(page.content, "html.parser")

    termCheck = soup.find_all("span", id="ctl00_ContentPlaceHolder1_lblHTML")
    termCheck = str(termCheck)
    if "Invalid Term" in termCheck:
        print("Error: " + term[0:len(term)-2] + " is invalid term.")
        print("Must be greater than 2023 and in format 20XX")
        print("If " + term[0:len(term)-2] + " is next semester's year, class schedules may not have been posted yet")
        time.sleep(5)
        sys.exit()

    webTime = soup.find_all("div", class_="col-sm-2 column-sm") #contains times and days
    tempWebProf = soup.find_all("div", class_="col-sm-4 column4-sm")
    tempWebProf = str(tempWebProf)

    webProf = []
    tempWebProf = tempWebProf.splitlines()
    for y in tempWebProf:
        if y.find("STAFF") != -1:
            webProf.append("Staff")
        else:
            temp = y[y.find(">") + 1:y.find("</a")]
            """50 is just to filter out garbage html code. Each line isn't always a name,
            but I don't know anyone with a name longer than 50 characters, so that makes only
            names to be included in webProf"""
            if len(temp) > 0 and len(temp) < 50 and temp[0] != " ":
                webProf.append(temp)


    webCourseName = soup.find_all("div",class_="col-md-10")

    webCourseNum = soup.find_all("div",class_="col-md-2")

    if webCourseNum == []:
        noCourseNumArr.append(prefix + " " + num)


    tempCRN = soup.find_all("div", class_="col-sm-1 hidden-xs")

    CRNs = []

    for i in range(0,len(tempCRN)):
        if i % 2 == 0:
            y = str(tempCRN[i])
            y = y[y.find(">")+1:y.find("</")]
            CRNs.append(y)

    tempSections = soup.find_all("div",class_="col-sm-1")
    tempSections = str(tempSections)

    tempSections = BeautifulSoup(tempSections, "html.parser")

    tempSections = tempSections.find_all("div",string=None)

    tempSections = [div for div in tempSections if div.find('strong')]

    sections = []
    tempSections = str(tempSections)
    tempSections = tempSections.splitlines()
    tempSections = tempSections[1:]

    for y in tempSections:
        y = y[y.find("<strong>") + 8:y.find("</strong>")]
        if len(y) < 3:
            sections.append(y)

    webReqSection = soup.find_all("div", class_="col-sm-offset-4 col-sm-8")
    tempReqSection = []
    for y in webReqSection:
        y = str(y)
        if "Must enroll in Section" in y:
            tempReqSection.append(y)

    reqSections = []
    for y in tempReqSection:
        y = y[y.find("<br/>") - 2:y.find("<br/>")]
        reqSections.append(y)
    if len(reqSections) > 0:
        reqSectionClasses.append(x)

    webWaitlist = soup.find_all("div", class_="col-sm-1 column-lg hidden-xs")

    webSeats = soup.find_all("div", class_="col-sm-1")
    webSeats = str(webSeats)
    webSeats = webSeats.split(", ")

    tempSeats = []
    for y in webSeats:
        if "Seats Avail:" in y:
            tempSeats.append(y)


    """Scrapes webpage for important info on each class and puts them into an array of classes of type course"""
    for i in range(0,len(sections)):
        courseName = str(webCourseName)
        courseName = courseName[courseName.find("strong")+7:courseName.find("</strong>")]
        if courseName.find("&amp;") > 0:
            courseName = courseName.replace("&amp;", "&")

        courseNum = str(webCourseNum)
        courseNum = courseNum[courseNum.find("strong") + 7:courseNum.find("</strong>")]

        seats = tempSeats[i]
        seats = seats[len(seats)-12:len(seats)-6]
        seats = int(seats)
        if seats < 0:
            seats = 0
        seats = str(seats)

        tempWaitlist = webWaitlist[i]
        tempWaitlist = str(tempWaitlist)
        waitlist = tempWaitlist[tempWaitlist.find(">")+1:tempWaitlist.find("</")]
        waitlist = waitlist.strip()

        times = []
        days = []
        temp = webTime[i].find('td')
        temp = str(temp)


        temp = temp[88:110]
        temp = temp.lstrip()
        temp = temp.rstrip()
        if temp != '':
            if temp.find('<br/>') > 0:
                days.append(temp[0:temp.find('<br/>')])
                days.append(temp[temp.find('>')+1:])
            else:
                days.append(temp[0:2])
                days[0].rstrip()
        else:
            days.append("Online")

        niceTime = []

        temp = webTime[i].find('tr')
        temp = str(temp)


        temp = temp[215:262]

        temp = temp.lstrip()
        temp = temp.rstrip()


        if temp != '':
            if temp.find('<br/>') > 0:
                niceTime.append(temp[0:15])
                niceTime.append(temp[20:])
            else:
                niceTime.append(temp[0:15])
            times = convertToMilitaryTime(niceTime)

            for j in range(0,len(niceTime)):
                tempTime = niceTime[j]
                tempTime = tempTime[0:2] + ":" + tempTime[2:11] + ":" + tempTime[11:]

                if tempTime[10] == '0':
                    tempTime = tempTime[0:10] + tempTime[11:]
                if tempTime[0] == '0':
                    tempTime = tempTime[1:]
                niceTime[j] = tempTime
        else:
            times.append([-1,-1])
            niceTime = "Online"
        prof = webProf[i]
        section = sections[i]
        if str(section).isalpha():
            courseName = courseName + " (lab)"
        CRN = CRNs[i]
        temp = Course(courseNum,courseName,prof,days,times,niceTime,section,CRN,seats,waitlist)
        #temp.printCourseInfo()
        classes.append(temp)
    #sys.exit()
#############################################################################################################
for y in noCourseNumArr:
    print("No course with name \"" + y + "\" found.")

if len(noCourseNumArr) > 0:
    time.sleep(5)
    sys.exit(0)


reqClasses = []

for x in reqClassesArr:
    classSections = []
    for y in classes:
        if x == y.course_num:
            #classSections.append(y) uncomment and delete if and else below if problem with adding classes
            if y.course_num in reqPreferredProfClass:
                if y.prof in reqPreferredProf:
                    classSections.append(y)
            else:
                classSections.append(y)
    reqClasses.append(classSections)

creditSum = 0
for x in reqClasses:
    creditSum += x[0].hours
if creditSum > maxHours:
    print("Error: Required classes exceed maximum desired credit hours")
    print("Please enter valid required classes or increase maximum desired hours")
    time.sleep(5)
    sys.exit()

print("Sorting...")

electiveClasses = []


for x in electiveClassesArr:
    classSections = []
    for y in classes:
        if x == y.course_num:
            #classSections.append(y) uncomment and delete if and else below if problem with adding classes
            if y.course_num in electivePreferredProfClass:
                if y.prof in electivePreferredProf:
                    y.elective = True
                    classSections.append(y)
            else:
                y.elective = True
                classSections.append(y)
    electiveClasses.append(classSections)

validSchedules = Schedules()
baseHours = 0

for r in reqClasses:
    baseHours += r[0].hours

freeHours = maxHours - baseHours

allClasses = []
for x in reqClasses:
    for y in x:
        allClasses.append(y)

for x in electiveClasses:
    for y in x:
        allClasses.append(y)

#############################################################################################################
"""Finds the maximum number of classes you can take as electives and still be under
the user set limit of credit hours per semester"""
maximizedNumClasses = 0
if electiveClasses[0] != []:
    creditHoursArr = []
    for x in electiveClasses:
        creditHoursArr.append(x[0].hours)

    creditHoursArr = sorted(creditHoursArr)
    tempHours = freeHours
    while tempHours > 0 and maximizedNumClasses < len(electiveClasses):
        tempHours -= creditHoursArr[maximizedNumClasses]
        maximizedNumClasses += 1

"""Checks if class is a lab"""
weirdLab = False
tempWeirdLab = False
for x in electiveClasses:
    #weirdLab = False
    for y in x:
        if y.isLab:
            weirdLab = True
            tempWeirdLab = True

if tempWeirdLab:
    maximizedNumClasses += 1
tempWeirdLab = False

numReqClasses = len(reqClasses)

for x in reqClasses:
    #weirdLab = False
    for y in x:
        if y.isLab:
            weirdLab = True
            tempWeirdLab = True
if tempWeirdLab:
    numReqClasses += 1

numberedClasses = [i for i in range(0, len(allClasses))]

choose = numReqClasses + maximizedNumClasses

print("Calculating combinations...")
combs = combinations(numberedClasses, choose)

#total = math.comb(len(numberedClasses), choose)

profFlagCount = []
for i in range(0,len(preferredProfClass)):
    profFlagCount.append(0)

#############################################################################################################

"""Filters all combinations within user parameters and time constraints"""
print("Comparing combinations...")
for x in combs:
    tempSchedule = []
    timeConflict = False
    if weirdLab:
        labWithClass = False
    else:
        labWithClass = True
    for y in x:
        tempSchedule.append(allClasses[y])

    profFlag = False
    profIndex = -1
    for i in range(0, len(preferredProfClass)):
        for j in tempSchedule:
            if j.course_num == preferredProfClass[i]:
                if preferredProf[i].lower() not in j.prof.lower():
                    profFlag = True
                    break
                else:
                    profFlagCount[i] += 1
        if profFlag:
            break

    if profFlag:
        continue

    for i in range(0,len(tempSchedule)):

        class1 = tempSchedule[i]
        for j in range(i+1,len(tempSchedule)):
            sameDay = False
            class2 = tempSchedule[j]
            for l in class1.days:
                for k in class2.days:
                    for q in l:
                        for w in k:
                            if q == w:
                                sameDay = True

            """If there is a weird lab in the schedule, checks if the class has concurrent lab"""
            if weirdLab and class1.course_num == class2.course_num and (class1.isLab ^ class2.isLab):
                labWithClass = True

            """Checks if classes are the same course"""
            if class1.course_num == class2.course_num and not class1.isLab and not class2.isLab:
                timeConflict = True
                break

            """Checks if classes are the same lab"""
            if class1.course_num == class2.course_num and class1.isLab and class2.isLab:

                timeConflict = True
                break

            """Checks for time conflicts"""
            if class1.timeEnd > class2.timeBegin and class2.timeEnd > class1.timeBegin and sameDay and class1.course_num != class2.course_num and class1.niceTime != "Online" and class2.niceTime != "Online":
                timeConflict = True
                break

            """Checks if class has lab with required section"""
            if class1.course_num == class2.course_num and class1.course_num in reqSectionClasses:
                if "(lab)" in class1.course_name:
                    temp1 = letterToNum(class1.section)
                else:
                    temp1 = int(class1.section)

                if "(lab)" in class2.course_name:
                    temp2 = letterToNum(class2.section)
                else:
                    temp2 = int(class2.section)
                if temp1 != temp2:
                    timeConflict = True
                    break



    if not timeConflict and labWithClass:
        validSchedules.addSchedule(tempSchedule)

noProfFlag = False
for i in range(0,len(profFlagCount)):
    if profFlagCount[i] == 0:
        print("Error: No professor with name \"" + preferredProf[i] + "\" found for " + preferredProfClass[i] + ".")
        noProfFlag = True
if noProfFlag:
    time.sleep(5)
    sys.exit()

print("Filtering results...")
filteredSchedules = []
for x in validSchedules.schedule_list:
    classCount = 0
    for y in x:
        for z in reqClassesArr:
            if y.course_num == z:
                classCount += 1
    if classCount == numReqClasses:
        filteredSchedules.append(x)

print()
print("Complete!")
print()
validSchedules.schedule_list = filteredSchedules
etime = time.time()

dt = etime - stime
#print(dt)

#validSchedules.printSchedule()
numValidSchedules = validSchedules.getNumSchedules()
print("Valid Schedules: " + str(numValidSchedules))

fileName = "SchedulesTXT.txt"

if saveToTxt and numValidSchedules > 0:
    i = 1
    while os.path.exists(fileName):
        fileName = "SchedulesTXT (" + str(i) + ").txt"
        i = i + 1
    file = open(fileName, "w")
    validSchedules.writeToFile(file)
    file.close()

    print("Schedules successfully exported to \"" + fileName + "\"!")

i = 0
tempData = validSchedules.getData()
data = []
for x in tempData:
    i = i + 1
    data.append([makeHeader(i)])
    for y in x:
        data.append([y])
    data.append([{"Course Number": ""}])



fileName = "SchedulesXLSX.xlsx"

if saveToExcel and numValidSchedules > 0:
    if os.path.exists(fileName):
        i = 1
        while os.path.exists(fileName):
            fileName = "SchedulesXLSX (" + str(i) + ").xlsx"
            i = i + 1

    with pd.ExcelWriter(fileName, engine="openpyxl", mode="w") as writer:

        for i, chunk in enumerate(data):
            df = pd.DataFrame(chunk)

            df.to_excel(writer, index=False, header=(i == 0), startrow=writer.sheets["Sheet1"].max_row if i > 0 else 0)

    print("Schedules successfully exported to \"" + fileName + "\"!")


time.sleep(5)

