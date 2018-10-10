# Outlook Room Finder
Innovative Ways to Find &amp; Book Available Meeting Rooms in any organization

Here you can find VBA code to be used with Outlook 2016/Office 365 .   

Please follow these steps to get it running :

1) Download the files and keep at one location
2) Modify following lines in the code as per your requirement 
```
    MAXROOMS = 44 ' Number of total meeting rooms
    MAXLOCATIONS = 4   'Total number of floors
    MAXROOMTYPES = 4   'Total groups of rooms 
    
    
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
 ```
 3) Go to your outlook and press `Alt + F11`
 
 ## <<  IN PROGRESS  >> 