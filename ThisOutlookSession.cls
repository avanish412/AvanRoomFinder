VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Global space
   Dim myOlApp As New Outlook.Application
   Dim myNameSpace As Outlook.NameSpace

Public Sub ShowMyForm()
  frmBookRoom.Show vbModeless
  
  'Initialize
  Set myNameSpace = myOlApp.GetNamespace("MAPI")
End Sub

Function GetFreeBusyInfo(strResourceId As String, datStart As Date, lngLength As Long) As Boolean
   
   Dim myRecipient As Outlook.Recipient
   Dim myFBInfo As String
   Dim datTemp As Date
   Dim lngSlot As Long
   
   Set myRecipient = myNameSpace.CreateRecipient(strResourceId)
   On Error GoTo ErrorHandler
   myFBInfo = myRecipient.FreeBusy(datStart, lngLength)
   
   datTemp = Month(datStart) & "/" & Day(datStart) & "/" & Year(datStart) & " 12:00:00 AM"
    lngSlot = (DateDiff("n", datTemp, datStart) / lngLength) + 1
    GetFreeBusyInfo = IIf(Mid(myFBInfo, lngSlot, 1) = "0", True, False)

   Exit Function
ErrorHandler:
    'MsgBox "Cannot access the information for " & strPerson
End Function

Function GetFreeBusyInfo2(strResourceId As String, datStart As Date, lngLength As Long, ByRef AvailableTime As Date) As Boolean

   Dim myRecipient As Outlook.Recipient
   Dim myFBInfo As String
   Dim lngSlot As Long
   Dim TargetDuration As String
   Dim StringtoSearch As String
   Dim datTemp As Date
   Dim startIndex As Long
   
   Set myRecipient = myNameSpace.CreateRecipient(strResourceId)
   On Error GoTo ErrorHandler
   myFBInfo = myRecipient.FreeBusy(datStart, 30)
   
   datTemp = Month(datStart) & "/" & Day(datStart) & "/" & Year(datStart) & " 12:00:00 AM"
   startIndex = (DateDiff("n", datTemp, datStart) / 30) + 1
   
   '''Target String
   TargetDuration = Mid(myFBInfo, startIndex, 36 - startIndex + 1)
   For i = 0 To lngLength / 30 - 1
    StringtoSearch = StringtoSearch & "0"
   Next i
    
   Dim position As Integer
   position = -1
   position = InStr(TargetDuration, StringtoSearch)
   
   'Find time
   If (position > 0) Then
        AvailableTime = Month(datStart) & "/" & Day(datStart) & "/" & Year(datStart) & " 12:00:00 AM"
        AvailableTime = AvailableTime + TimeSerial(0, 30 * (startIndex - 1 + position - 1), 0)
        GetFreeBusyInfo2 = True
    Else
        GetFreeBusyInfo2 = False
    End If
    
   Exit Function
ErrorHandler:
    'MsgBox "Cannot access the information for " & strPerson
End Function

Function GetFreeBusyInfo3(strResourceId As String, datStart As Date, lngLength As Integer) As Integer

   Dim myRecipient As Outlook.Recipient
   Dim myFBInfo As String
   Dim datTemp As Date
   Dim startIndex As Long
   Dim AvailableDuration As Integer
   
   AvailableDuration = 0
   
   Set myRecipient = myNameSpace.CreateRecipient(strResourceId)
   On Error GoTo ErrorHandler
   myFBInfo = myRecipient.FreeBusy(datStart, 30)
   
   
   datTemp = Month(datStart) & "/" & Day(datStart) & "/" & Year(datStart) & " 12:00:00 AM"
   startIndex = (DateDiff("n", datTemp, datStart) / 30) + 1
    
    Do While Mid(myFBInfo, startIndex, 1) = "0"
        AvailableDuration = AvailableDuration + 30
        startIndex = startIndex + 1
    Loop

   GetFreeBusyInfo3 = AvailableDuration
   
   Exit Function
ErrorHandler:
    'MsgBox "Cannot access the information for " & strPerson
End Function

Function IsFreeXXXX(strPerson As String, datStart As Date, lngLength As Long) As Boolean
    Dim olkRecipient As Outlook.Recipient, _
        olkAddress As Outlook.AddressEntry, _
        olkUser As Outlook.ExchangeUser, _
        strFreeBusy As String, _
        lngSlot As Long, _
        datTemp As Date
    Set olkRecipient = Outlook.Session.CreateRecipient(strPerson)
    If TypeName(olkRecipient) <> "Nothing" Then
        Set olkAddress = olkRecipient.AddressEntry
        If olkAddress.AddressEntryUserType = olExchangeUserAddressEntry Then
            Set olkUser = olkAddress.GetExchangeUser
            strFreeBusy = olkUser.GetFreeBusy(datStart, lngLength)
            datTemp = Month(datStart) & "/" & Day(datStart) & "/" & Year(datStart) & " 12:00:00 AM"
            lngSlot = (DateDiff("n", datTemp, datStart) / lngLength) + 1
            IsFree = IIf(Mid(strFreeBusy, lngSlot, 1) = "0", True, False)
        End If
    End If
    Set olkUser = Nothing
    Set olkAddress = Nothing
    Set olkRecipient = Nothing
End Function

Function CreateAppt(strResourceName As String, strResourceId As String, datStart As Date, lngLength As Long, subject As String)
 Dim myItem As Object
 Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient
   
 Set myItem = Application.CreateItem(olAppointmentItem)
 myItem.MeetingStatus = olMeeting
 myItem.subject = subject
 myItem.Location = strResourceName
 myItem.Start = datStart
 myItem.duration = lngLength
 Set myRequiredAttendee = myItem.Recipients.Add(strResourceId)
 myRequiredAttendee.Type = olRequired
 Set myResourceAttendee = myItem.Recipients.Add(strResourceId)
 myResourceAttendee.Type = olResource
 
 'myItem.Display
 myItem.Send
End Function
