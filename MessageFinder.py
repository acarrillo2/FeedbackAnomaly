import win32com
import win32com.client
import string
import os
# the findFolder function takes the folder you're looking for as folderName,
# and tries to find it with the MAPIFolder object searchIn
# Source: https://stackoverflow.com/questions/2043980/using-win32com-and-or-active-directory-how-can-i-access-an-email-folder-by-name
# Modified by acarrillo2 2018-04-08

def findFolder(folderName,searchIn):
    try:
        lowerAccount = searchIn.Folders
        for x in lowerAccount:
            if x.Name == folderName:
##                print('found it %s'%x.Name)
                objective = x
                return objective
        return None
    except Exception as error:
        print("Looks like we had an issue accessing the searchIn object")
        print(error)
        return None

def MessageFinder(email, folder1Name, folder2Name):

    outlook=win32com.client.Dispatch("Outlook.Application")

    ons = outlook.GetNamespace("MAPI")

    #this is the initial object you're accessing, IE if you want to access
    #the account the Inbox belongs too
    one = email

    #Retrieves a MAPIFolder object for your account 
    #Object functions and properties defined by MSDN at 
    #https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mapifolder_members(v=office.14).aspx
    Folder1 = findFolder(one,ons)

    #Now pass you're MAPIFolder object to the same function along with the folder you're searching for
    Folder2 = findFolder(folder1Name,Folder1)

    #Rinse and repeat until you have an object for the folder you're interested in
    Folder3 = findFolder(folder2Name,Folder2)

    #This call returns a list of mailItem objects refering to all of the mailitems(messages) in the specified MAPIFolder
    messages = Folder3.Items

    return messages

