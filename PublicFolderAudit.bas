Sub GetPublicFoldersWithItemCountsAndLastReceivedDate()
    Dim olNamespace As Outlook.NameSpace
    Dim olRootFolder As Outlook.Folder
    Dim output As String
    Dim folderCount As Long
    Dim outputFilePath As String

    ' Initialize the output string with headers
    output = "Folder Path, Item Count, Last Received Date" & vbCrLf

    ' Get the Outlook Namespace
    Set olNamespace = Application.GetNamespace("MAPI")

    ' Get the root public folder (customize the folder selection as needed)
    Set olRootFolder = olNamespace.Folders.Item(1).Folders.Item(2) ' Adjust as necessary

    ' Initialize folder count
    folderCount = 0

    ' Process the root folder and its subfolders recursively
    Call ProcessFolder(olRootFolder, output, folderCount, 0)

    ' Specify the output CSV file path
    outputFilePath = "C:\outputpath.csv" ' Replace with your desired file path

    ' Write the results to the CSV file
    Open outputFilePath For Output As #1
    Print #1, output
    Close #1

    MsgBox "Data export complete! " & folderCount & " folders processed.", vbInformation
End Sub

Sub ProcessFolder(ByVal folder As Outlook.Folder, ByRef output As String, ByRef folderCount As Long, ByVal level As Integer)
    Dim olItems As Outlook.Items
    Dim item As Object
    Dim latestReceivedDate As Variant
    Dim folderPath As String
    Dim itemCount As Long
    Dim subFolder As Outlook.Folder
    Dim hierarchyPrefix As String

    ' Increment the folder count
    folderCount = folderCount + 1

    ' Create the hierarchy prefix based on the level
    hierarchyPrefix = String(level * 3, "-->") ' Use '|->' for each level

    ' Combine the hierarchy prefix and folder name
    folderPath = hierarchyPrefix & " " & folder.Name

    ' Escape commas in the folder name
    If InStr(folderPath, ",") > 0 Then
        folderPath = """" & folderPath & """" ' Escape commas by enclosing in quotes
    End If

    ' Get items in the folder and sort by ReceivedTime (latest first)
    Set olItems = folder.Items
    olItems.Sort "[ReceivedTime]", True

    ' Initialize the latest received date
    If olItems.Count > 0 Then
        Set item = olItems.Item(1)
        If TypeOf item Is Outlook.MailItem Then
            latestReceivedDate = item.ReceivedTime
        Else
            latestReceivedDate = "N/A"
        End If
    Else
        latestReceivedDate = "No Emails"
    End If

    ' Count the number of items in the folder
    itemCount = olItems.Count

    ' Append folder info to the output string
    output = output & folderPath & ", " & itemCount & ", " & latestReceivedDate & vbCrLf

    ' Recursively process each subfolder, increasing the level
    For Each subFolder In folder.Folders
        Call ProcessFolder(subFolder, output, folderCount, level + 1)
    Next subFolder
End Sub
