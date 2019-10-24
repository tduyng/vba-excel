Public WithEvents GItems As Outlook.Items
'Updated by ExtendOffice 20180814
Private Sub Application_Startup()
    Set GItems = Outlook.Application.Session.GetDefaultFolder(olFolderInbox).Items
End Sub
Private Sub GItems_ItemAdd(ByVal Item As Object)
Dim xMtRequest As MeetingItem
Dim xAppointmentItem As AppointmentItem
Dim xMtResponse As MeetingItem
If Item.Class = olMeetingRequest Then
    Set xMtRequest = Item
    Set xAppointmentItem = xMtRequest.GetAssociatedAppointment(True)
    If xAppointmentItem.GetOrganizer.Name = "Sender Name" Then
        With xAppointmentItem
            .ReminderMinutesBeforeStart = 45
            .Categories = "Orange Category"
            .Save
        End With
        Set xMtResponse = xAppointmentItem.Respond(olMeetingAccepted)
        xMtResponse.Send
        xMtRequest.Delete
    End If
End If
End Sub



Sub AcceptMeeting() 
 Dim myNameSpace As Outlook.NameSpace 
 Dim myFolder As Outlook.Folder 
 Dim myMtgReq As Outlook.MeetingItem 
 Dim myAppt As Outlook.AppointmentItem 
 Dim myMtg As Outlook.MeetingItem 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox) 
 Set myMtgReq = myFolder.Items.Find("[MessageClass] = 'IPM.Schedule.Meeting.Request'") 
 If TypeName(myMtgReq) <> "Nothing" Then 
 Set myAppt = myMtgReq.GetAssociatedAppointment(True) 
 Set myMtg = myAppt.Respond(olResponseAccepted, True) 
 myMtg.Send 
 End If 
End Sub