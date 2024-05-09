VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private objNS As Outlook.NameSpace
Private WithEvents objNewMailItems As Outlook.Items
Attribute objNewMailItems.VB_VarHelpID = -1

Private Sub Application_Startup()
 
Dim objWatchFolder As Outlook.Folder
Set objNS = Application.GetNamespace("MAPI")

'Set the folder and items to watch:
Set objWatchFolder = objNS.GetDefaultFolder(olFolderInbox)
Set objNewMailItems = objWatchFolder.Items

Set objWatchFolder = Nothing
End Sub

Private Sub objNewMailItems_ItemAdd(ByVal item As Object)
Dim objMeetingItem As Outlook.meetingItem

If item.Class <> olMeetingCancellation Then Exit Sub
Set objMeetingItem = item

AutoDeleteMeeting objMeetingItem

Set item = Nothing
End Sub


'perform the guts of the calendar cleanup
Private Sub AutoDeleteMeeting(ByVal item As Object)

Dim oMeetingItem As Outlook.meetingItem

Dim oAppointmentItem As Outlook.AppointmentItem
    
    Set oMeetingItem = item

    If Len(oMeetingItem.Subject) > 0 And InStr(1, oMeetingItem.Subject, "Canceled:") <> 0 Then
        LogMessage ("AutoDeleteMeeting():" & oMeetingItem.Subject)
        
        Set oAppointmentItem = oMeetingItem.GetAssociatedAppointment(False)
        
        If Not oAppointmentItem Is Nothing Then
            LogMessage ("AutoDeleteMeeting():" & oMeetingItem.Subject & " @ " & oAppointmentItem.Start)
            oAppointmentItem.Delete
        Else
            LogMessage (oMeetingItem.Subject & ": can't find associated appointment")
        End If
       oMeetingItem.Delete
    Else
      LogMessage ("AutoDeleteMeeting(): Subject doesn't match pattern")
       
    End If
End Sub

'perform the guts of the calendar cleanup
Public Sub OpenAssociatedMeeting(oMeetingItem As meetingItem)

Dim oAppointmentItem As AppointmentItem
    
        Set oAppointmentItem = oMeetingItem.GetAssociatedAppointment(False)
        
        If Not oAppointmentItem Is Nothing Then
            LogMessage (oMeetingItem.Subject & " @ " & oAppointmentItem.Start)
            oMeetingItem.Open
        Else
            LogMessage (oMeetingItem.Subject & ": can't find associated appointment")
            MsgBox ("Can't find associated meeting")
        End If
    
End Sub

Sub LogMessage(message As String)
'PURPOSE: Add More Text To The End Of A Text File
'SOURCE: www.TheSpreadsheetGuru.com

Dim TextFile As Integer
Dim FilePath As String
Dim timestamp As String
timestamp = Format(Now(), "yyyy-MM-dd hh:mm:ss")

Debug.Print timestamp & ": " & message

'What is the file path and name for the new text file?
  FilePath = Environ$("USERPROFILE") & "\outlooklog.txt"

'Determine the next file number available for use by the FileOpen function
  TextFile = FreeFile

'Open the text file
  Open FilePath For Append As TextFile

'Write some lines of text
  Print #TextFile, timestamp & ": " & message
  
'Save & Close Text File
  Close TextFile

End Sub

'this can be called from the Macros menu to perform the delete function on each selected meeting cancellation
Sub DeleteMeeting()
    Dim x As Object
    For Each x In Application.ActiveExplorer.Selection
        If TypeName(x) = "MeetingItem" Then
            Dim meetingItem As Outlook.meetingItem
            Set meetingItem = x
            Call AutoDeleteMeeting(meetingItem)
        End If
    Next
End Sub



Public Sub DeleteSelectedMeetings()
    Dim objOL As Outlook.Application
    Dim currentExplorer As Explorer
    Dim Selection As Selection
    
    Dim obj As Object
    Set objOL = Outlook.Application
    Set currentExplorer = objOL.ActiveExplorer
    Set Selection = currentExplorer.Selection

    For Each obj In Selection
 
     With obj
 
       If TypeName(obj) = "MeetingItem" Then
           Debug.Print .Subject
           Dim meetingItem As Outlook.meetingItem
           Set meetingItem = obj
           Call AutoDeleteMeeting(meetingItem)
       End If

     End With

    Next

    Set Selection = Nothing
    Set currentExplorer = Nothing
    Set obj = Nothing
    Set Selection = Nothing

End Sub



Public Sub ApptFromMail(item As MailItem)

Dim objAppt As Outlook.AppointmentItem
Set objAppt = Application.CreateItem(olAppointmentItem)

With objAppt
    .Subject = item.Subject
    .Location = "Concur"
    .AllDayEvent = False
    .BusyStatus = olFree
    .ReminderSet = True
    .ReminderMinutesBeforeStart = 0
    
    .Start = Date + 3
    .Body = item.Body
    .Save

End With
    Set objAppt = Nothing
End Sub



