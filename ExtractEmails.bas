'----------------------------------------------
' Author: Reynaldo Cortez
' Date: 2023-07-19
' Description: This script extracts email addresses from the 
' selected Outlook folder and exports them to a text file.
'
' Please ensure that you have access permissions to write a file to "C:\EmailAddresses.txt", or change this path to a suitable location where you have write permissions.
'----------------------------------------------

' Include a reference to Microsoft Scripting Runtime by checking it in the "Tools -> References" list
' This is needed for the FileSystemObject

' Declare the Sub procedure
Sub ExtractEmails()

    ' Declare the variables
    Dim objNS As Namespace
    Dim objFolder As MAPIFolder
    Dim objDict As Object
    Dim objItem As Object
    Dim objMail As MailItem
    Dim objFSO As Object
    Dim objTextFile As Object
    Dim strEmail As String
    Dim varEmail As Variant

    ' Set a reference to the namespace object
    Set objNS = Application.GetNamespace("MAPI")
    
    ' Allow the user to select a folder in Outlook
    Set objFolder = objNS.PickFolder
    
    ' Create an instance of the Dictionary object
    Set objDict = CreateObject("Scripting.Dictionary")
    
    ' Loop through each item in the folder
    For Each objItem In objFolder.Items
        
        ' Only continue if the item is a MailItem
        If TypeName(objItem) = "MailItem" Then
            
            ' Cast the item to a MailItem object
            Set objMail = objItem
            
            ' Extract the email address from the SenderEmailAddress property
            strEmail = objMail.SenderEmailAddress
            
            ' Add the email address to the Dictionary object (to avoid duplicates)
            If Not objDict.Exists(strEmail) Then
                objDict.Add strEmail, ""
            End If
            
        End If
    Next objItem
    
    ' Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Create a new text file to save the email addresses
    Set objTextFile = objFSO.CreateTextFile("C:\EmailAddresses.txt", True)
    
    ' Loop through each key (email address) in the Dictionary object
    For Each varEmail In objDict.Keys
        
        ' Write the email address to the text file
        objTextFile.WriteLine varEmail
    Next varEmail
    
    ' Close the text file
    objTextFile.Close
    
    ' Clean up
    Set objTextFile = Nothing
    Set objFSO = Nothing
    Set objMail = Nothing
    Set objItem = Nothing
    Set objDict = Nothing
    Set objFolder = Nothing
    Set objNS = Nothing

    ' Inform the user that the email extraction is complete
    MsgBox "The email extraction is complete. The email addresses have been saved to C:\EmailAddresses.txt.", vbInformation

End Sub
