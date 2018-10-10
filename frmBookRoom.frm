VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBookRoom 
   Caption         =   "Book My Meeting Room"
   ClientHeight    =   5390
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6710
   OleObjectBlob   =   "frmBookRoom.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBookRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Global Variables
Dim rooms(0 To 3, 0 To 43) As String
Dim MAXROOMS As Long
Dim MAXLOCATIONS As Long
Dim Location(0 To 3) As String
Dim RoomType(0 To 3) As String


Private Sub btnAvailableRooms_Click()
    
    Dim startDate As Date
    Dim duration As Long
    Dim selectedLocation As String
    lbAvailableRooms.Clear
    
    startDate = cmbDate.Value & " " & cmbTime.Value
    duration = cmbDuration.Value
    selectedLocation = cmbLocation.Value
    
    For i = 0 To MAXROOMS - 1
        If chkAllfloors.Value = True Or rooms(0, i) = selectedLocation Then
            If (chkSmallRoom.Value = True And rooms(3, i) = RoomType(0)) Or _
               (chkMediumRoom.Value = True And rooms(3, i) = RoomType(1)) Or _
               (chkLargeRoom.Value = True And rooms(3, i) = RoomType(2)) Then

                If ThisOutlookSession.GetFreeBusyInfo(rooms(2, i), startDate, duration) Then
                    lbAvailableRooms.AddItem (rooms(1, i))
                End If
            End If
        End If
    Next i
    
End Sub

Private Sub btnAvailableRooms2_Click()
    Dim startDate As Date
    Dim enddate As Date
    Dim duration As Long
    Dim selectedLocation As String
    
    lbAvailableRooms2.Clear
    
    startDate = cmbDate2.Value
    duration = cmbDuration2.Value
    selectedLocation = cmbLocation.Value
    enddate = startDate & " 6:00:00 PM"
    
    For i = 0 To MAXROOMS - 1
        startDate = cmbDate2.Value & " 10:00:00 AM"
        If chkAllfloors.Value = True Or rooms(0, i) = selectedLocation Then
            If (chkSmallRoom.Value = True And rooms(3, i) = RoomType(0)) Or _
               (chkMediumRoom.Value = True And rooms(3, i) = RoomType(1)) Or _
               (chkLargeRoom.Value = True And rooms(3, i) = RoomType(2)) Then
                
                Dim AvailableTime As Date
                
                Do While startDate < enddate
                    If ThisOutlookSession.GetFreeBusyInfo2(rooms(2, i), startDate, duration, AvailableTime) Then
                        lbAvailableRooms2.AddItem (rooms(1, i) & "@" & Format(AvailableTime, "h:mm:ss AM/PM"))
                        startDate = AvailableTime + TimeSerial(0, duration, 0)
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
    Next i
End Sub

Private Sub btnAvailableRooms3_Click()
    Dim startDate As Date
    Dim selectedduration As Integer
    Dim selectedLocation As String
    Dim remainingduration As Integer
    Dim TotalAvailableDuration As Integer
    
    lbAvailableRooms3.Clear
    
    TotalAvailableDuration = 0
    
    startDate = cmbDate3.Value & " " & cmbTime3.Value
    selectedduration = cmbDuration3.Value
    selectedLocation = cmbLocation.Value
    remainingduration = selectedduration
    
    Do While remainingduration <> 0
        Dim MaxAvailableDuration As Integer
        Dim maxIndex As Integer
        MaxAvailableDuration = 0: maxIndex = 0
        
        For i = 0 To MAXROOMS - 1
            If chkAllfloors.Value = True Or rooms(0, i) = selectedLocation Then
                If (chkSmallRoom.Value = True And rooms(3, i) = RoomType(0)) Or _
               (chkMediumRoom.Value = True And rooms(3, i) = RoomType(1)) Or _
               (chkLargeRoom.Value = True And rooms(3, i) = RoomType(2)) Then
               
                    Dim duration As Integer
                    ''''' Find maximum available slot in each time
                    duration = ThisOutlookSession.GetFreeBusyInfo3(rooms(2, i), startDate, selectedduration - MaxAvailableDuration)
                    If duration >= remainingduration Then
                        MaxAvailableDuration = remainingduration
                        maxIndex = i
                        Exit For
                    ElseIf duration > MaxAvailableDuration Then
                        MaxAvailableDuration = duration
                        maxIndex = i
                    End If
                End If
            End If
        Next i
        If MaxAvailableDuration > 0 Then
            TotalAvailableDuration = TotalAvailableDuration + MaxAvailableDuration
            
            lbAvailableRooms3.AddItem (rooms(1, maxIndex) & "@" & Format(startDate, "h:mm:ss AM/PM") & "@" & MaxAvailableDuration)
            
            startDate = startDate + TimeSerial(0, MaxAvailableDuration, 0)
            
            remainingduration = remainingduration - MaxAvailableDuration
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub btnBookRoom_Click()
    Dim selectedroom As String
    Dim startDate As Date
    Dim duration As Long
    
    Dim selectedroomid As String
    
    startDate = cmbDate.Value & " " & cmbTime.Value
    duration = cmbDuration.Value
    
    selectedroom = lbAvailableRooms.Value
    
    For i = 0 To MAXROOMS - 1
        If selectedroom = rooms(1, i) Then
            selectedroomid = rooms(2, i)
            Exit For
        End If
    Next i
    Dim v As Variant
    v = ThisOutlookSession.CreateAppt(selectedroom, selectedroomid, startDate, duration, txtSubject.Text)
    'DoCmd.Close frmBookRoom, Me.Name
    
End Sub

Private Sub btnBookRoom2_Click()
    Dim selectedroom As String
    Dim startDate As Date
    Dim duration As Long
    
    If lbAvailableRooms2.ListCount = 0 Then
        Exit Sub
    End If
    
    Dim LArray() As String

    LArray = Split(lbAvailableRooms2.Value, "@")
    
    selectedroom = LArray(0)
    
    Dim selectedroomid As String
    
    startDate = cmbDate2.Value & " " & LArray(1)
    duration = cmbDuration2.Value
    
    
    For i = 0 To MAXROOMS - 1
        If selectedroom = rooms(1, i) Then
            selectedroomid = rooms(2, i)
            Exit For
        End If
    Next i
    Dim v As Variant
    v = ThisOutlookSession.CreateAppt(selectedroom, selectedroomid, startDate, duration, txtSubject.Text)
End Sub

Private Sub btnBookRoom3_Click()
    Dim selectedroom As String
    Dim startDate As Date
    Dim duration As Long
    
    If lbAvailableRooms3.ListCount = 0 Then
        Exit Sub
    End If
    
    For j = 0 To lbAvailableRooms3.ListCount - 1
        
        Dim LArray() As String
        LArray = Split(lbAvailableRooms3.List(j), "@")
        
        selectedroom = LArray(0)
    
        Dim selectedroomid As String
        
        startDate = cmbDate3.Value & " " & LArray(1)
        duration = LArray(2)
        
        For i = 0 To MAXROOMS - 1
        If selectedroom = rooms(1, i) Then
            selectedroomid = rooms(2, i)
            Exit For
        End If
        Next i
        Dim v As Variant
        v = ThisOutlookSession.CreateAppt(selectedroom, selectedroomid, startDate, duration, txtSubject.Text)

    Next j
    
End Sub

Private Sub chkAllfloors_Click()
    If chkAllfloors.Value = True Then
        cmbLocation.Enabled = False
    Else
        cmbLocation.Enabled = True
    End If
    
End Sub

Private Sub UserForm_Initialize()
    MAXROOMS = 44 ' Number of total meeting rooms
    MAXLOCATIONS = 4   'Total number of floors
    MAXROOMTYPES = 4   'Total groups of rooms 
    
    'Subject
    txtSubject.Text = "My Team Meeting"
    
    'Locations
    Location(0) = "Bangalore Gnd Floor"
    Location(1) = "Bangalore 2nd Floor"
    Location(2) = "Bangalore 4th Floor"
    Location(3) = "Bangalore 5th Floor"
    
    'Room Type
    RoomType(0) = "Small"  '4-6 Seater
    RoomType(1) = "Medium" '8-10 Seater
    RoomType(2) = "Large"  '12-14 Seater
    RoomType(3) = "Extra Large" '20 to 40 Seaters or any Amphiteatre
    
    
    For i = 0 To MAXLOCATIONS - 1
        cmbLocation.AddItem (Location(i))
    Next i
    cmbLocation.ListIndex = 0
    
    'Room capacity check boxes
    chkSmallRoom.Value = True

   'Create Room Array for your office location 
    rooms(0, 0) = Location(2)
    rooms(1, 0) = "AP-BLR-4F-Ganga-VC(4Seater)"
    rooms(2, 0) = "AP-BLR-4F-Ganga-VC@mycompany.com"
    rooms(3, 0) = RoomType(0)
    
    rooms(0, 1) = Location(2)
    rooms(1, 1) = "AP-BLR-4F-Yamuna-VC(4Seater)"
    rooms(2, 1) = "AP-BLR-4F-Yamuna-VC@mycompany.com"
    rooms(3, 1) = RoomType(0)
	
    '...  continue with 
    '...  COMPLETE
	'...  ALL THE INITIALIZATION
	'...  HERE
    '...  TILL
	
    rooms(0, 43) = Location(1)
    rooms(1, 43) = "AP-BLR-2F-Godavari"
    rooms(2, 43) = "Godavari@mycompany.com"
    rooms(3, 43) = RoomType(0)
    
    'Populate Date
    For i = 0 To 9
    Dim strDate As String
        strDate = Date + i
        cmbDate.AddItem (strDate)
        cmbDate2.AddItem (strDate)
        cmbDate3.AddItem (strDate)
        cmbDate4.AddItem (strDate)
    Next i
    cmbDate.ListIndex = 0: cmbDate2.ListIndex = 0: cmbDate3.ListIndex = 0: cmbDate4.ListIndex = 0
    
    'Populate Time
    Dim time As Date
    time = "12:00:00 AM"
    
    Do While TimeValue(time) < TimeValue("11.30:00 PM")
        Dim strTime As String
        strTime = Format(time, "h:mm:00 AM/PM")
        cmbTime.AddItem (strTime)
        cmbTime3.AddItem (strTime)
        cmbTime4.AddItem (strTime)
        time = time + TimeSerial(0, 30, 0)
    Loop
    
    Dim currtime As Date
    currtime = Now
    currtime = Format(currtime, "h:00:00 AM/PM")
    currtime = currtime + TimeSerial(0, 30, 0)
    
    cmbTime.Value = Format(currtime, "h:mm:00 AM/PM")
    cmbTime3.Value = Format(currtime, "h:mm:00 AM/PM")
    cmbTime4.Value = Format(currtime, "h:mm:00 AM/PM")
    
    'Populate Duration
    For k = 30 To 480 Step 30
        cmbDuration.AddItem (k)
        cmbDuration2.AddItem (k)
        cmbDuration3.AddItem (k)
        cmbDuration4.AddItem (k)
    Next k
    cmbDuration.ListIndex = 0: cmbDuration2.ListIndex = 0: cmbDuration3.ListIndex = 0: cmbDuration4.ListIndex = 0
    
    rdoWeekly.Value = True
    chkMonday.Value = True
    
End Sub
