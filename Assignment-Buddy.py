import os, time, os.path, zipfile, shutil

def formatPath(path):
    path = path.replace("\\", "/")
    return path


def getDownloadPathCore():
    """Returns the default downloads path for linux or windows"""
    if os.name == 'nt':
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
        return location
    else:
        return os.path.join(os.path.expanduser('~'), 'downloads')


def getDownloadsPath():
    return formatPath(getDownloadPathCore())


def createFolder(directory):
    '''Attempts to create a new folder with the given directory'''
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print('Error: Creating directory ') + directory


def moveFile(downloadsPath, course, assignmentNumber, oldFileName, newFileName):
    '''Attempts to move the downloaded assignmetn folder to it's course folder'''
    oldDirectory = downloadsPath + '/assignment/' + oldFileName
    newDirectory = "C:/" + course + "/Assignments/Assignment " + str(assignmentNumber) + "/" + newFileName

    try:
        os.rename(oldDirectory, newDirectory)
    except OSError:
        print('Error: Moving file ' + oldFileName)


def unzipFolder(newDirectory):
    zip_ref = zipfile.ZipFile(newDirectory + '.zip', 'r')
    zip_ref.extractall(newDirectory)
    zip_ref.close()


def groupTogetherDownloadedPictures(downloadsPath):
    '''Moves all the pictures from the downloaded folders into one folder containing all
       of the assignment pictures'''
    assignmentPath = downloadsPath + '/assignment'
    assignmentPathZip = assignmentPath + '.zip'
    unzipFolder(assignmentPath)
    os.remove(assignmentPathZip)

    i = 1
    while os.path.exists(assignmentPath + ' (' + str(i) + ')' + '.zip'):
        directory = assignmentPath + ' (' + str(i) + ')'
        unzipFolder(directory)
        for pic in os.listdir(directory):
            try:
                os.rename(directory + '/' + pic, assignmentPath + '/' + pic)
            except OSError:
                print('Error: Moving file ' + pic)

        os.remove(directory + '.zip')
        shutil.rmtree(directory)
        i += 1


def movePictures(downloadsPath, course, assignmentNumber, questions):
    '''Renames the assigment pictures and moves them into the course assignment folder'''
    groupTogetherDownloadedPictures(downloadsPath)

    folderDirectory = "C:/" + course + "/Assignments/Assignment " + str(assignmentNumber)
    createFolder(folderDirectory)

    currDirectory = downloadsPath + '/assignment' # where the downloaded pics are
    courseCode = course[-3] + course[-2] + course[-1]
    question = 0
    part = 1
    for pic in os.listdir(currDirectory):
        currQuestion = questions[question]

        if currQuestion == 1:
            fileName = courseCode + " - A" + str(assignmentNumber) + " - Q" + str(question + 1) + ".png"
            moveFile(downloadsPath, course, assignmentNumber, pic, fileName)
            question += 1
            part = 1
        else:
            fileName = courseCode + " - A" + str(assignmentNumber) + " - Q" + str(question + 1) +  "." + str(part) + ".png"
            moveFile(downloadsPath, course, assignmentNumber, pic, fileName)
            part += 1

            if part > currQuestion:
                part = 1
                question += 1

    shutil.rmtree('C:/Users/Ogi/Downloads/assignment')


instructions = input("Do you need instructions? Enter 1 for yes, 2 for no: ")

if instructions != "2":
    print('''
    1. Take pictures of all the pages of your assignment.
    2. Send yourself an email with all the pictures of the assignment attached. Make the subject
       of the email "Assignment", and make sure you attach all the pictures in order. If you have
       too many pictures to attach in one email, spread them out across multiple emails, and make
       sure the subject is "Assignment" for all the emails.
    3. Go on your laptop and go to your email and just download the attachments which you have sent.
       Make sure you download them in order of most recently delivered.
    4. Finally, run this program and fill in the correct information.y
    ''')

downloadsPath = getDownloadsPath()
course = input("Course name: ")
assignment = int(input("Assignment number: "))
numQuestions = int(input("Number of questions: "))
numPages = []
for question in range(numQuestions):
    pages = input("Number of pages for question " + str(question + 1) + ": ")
    numPages.append(int(pages))

movePictures(downloadsPath, course, assignment, numPages)

print("\nYour assignment has been successfully saved.")

exit = input()
