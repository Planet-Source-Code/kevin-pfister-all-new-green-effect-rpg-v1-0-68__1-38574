Important notes for Green Effect

The keys for Green Effect are:

The arrow keys : Movement
Space Bar : Read Message / Talk to people
A : Attack

The Map files and other files for the game are housed in the GE folder in the main Green Effect Folder

The OLF editor is not fully used yet, it will be in the next version

There is no monster pic so currently a picture of you chases you and tries to kill you

Made so Far:

A small part of the Map, and the first two towns and the buildings in it
A complete shopping engine
The Startup, (The future part)
Movement Engine
Caption Engine
The Tile Engine and the Alpha Blending of the Tiles
A lot of Tiles!
Day and Night
Death
Lights
Saving and Loading
A menu system
A tracking Engine and fighting engine

Modifications from last version

Now uses bitblt instead of Paintpicture (Large Speed Improvement)
Now uses getpixel and setpixelV instead of Point and Pset (Large Speed Improvement)
Rewrote Tile Engine (Speed and reliablity Improvement)
Added to the map
Completed the shopping Engine	
Players movement is by BitBlt instead of an image (No Flickering and speed improvement)
Random Messages now can be different depending on if it is day or night
Messages appear when space is pressed near certain objects eg. Bookcases, Beds...
Added comments
Able to find objects, not just castras (able to, but no objects added yet)
Totaly remade the PLS editor using extracts from the OLF Editor, also sorts characters for readability

Whats New

A tracking engine (for the monsters)
Saving and loading (Saving only works when outside of a building)
Added a menu system with options
Adding Light Sources (For Use at Night)
Added preview window to Map Editor
Made the intro
Added Save function to Character Creator now with mask save
A tile editing program
A fighting engine
Keys to Enter Buildings
When a door is added to the Map Editor, it updates the *.Hus file, creates a new map and makes a blank map 

Bugs Fixed
	
Alpha Blending of Tiles fixed, uses correct colours
Character Flickering, Changed to BitBlt
Character Creator not changing characters colours
Monster Character not shown until fully faded in
Monster disappears when player dies
	
To Do
	
Complete the Map
Create Enemy characters
Create a better storyline
Create a lot of new tiles
Create weapon pictures
Add more comments
Add Map like from Zelda
Add Mask Saving to Character Creator
Use the tileset instead of lots of pictureboxes
Ability to buy buildings (Estate agent is made but not coded)
Make monster also track by waypoints
Encrypt Maps, *.Hus file and *.PLS
	
Long Term
	
Turn into multiplayer where each group can buy buildings and play together as teams
Port to DirectX ,Once I download the SDK :)

Bugs to Fix

Lighting is not centered

What does not fully work

The New character graphics has no load functions
The Pls Editor does not allow for the creation of new characters
The Pls Editor has not been updated for the new feators of Green Effect

What Files Contain
	
*.Hus
This contains the Position of the Doors in the map, this file is only for the main map file
*.Map
This contains the Map and the tile information. It is in a Semi Compressed State 
There is one for the main map and one for each house
*.Pls
This contains the Position of the Extra Characters, it also contains what they say...
There is one for the main map and one for each house
*.OLF
This is the Object Location File it tells the program where the objects can be found such as Keys, Castras...

To make a new building

Load the map editor and load the GEffect.Map file
From this start making your house, when you add a door it will ask if you would like it to update the *.Hus file and create the template files, click Yes for both
Save the file and exit the map editor, the templates that have been created are a *.map file, a *.Pls file and a *.OLF file. 
Use the map editor to create the building, the entry points are 105,110 base your building around it(Remember that square has to be passable)
Save the map and open the Pls editor, with this you place the extra people which will talk to you click on the map where you want to place a person and follow the steps

How the Files Work

*.Hus
The Hus file contains the information about the buildings, like where they are what they are called which key is needed. There is only 1 Hus file GEffect.Hus which is for the main map file

Eg..

1	<-- This is the number of Houses in the File 
##Name House	<-- The Name of the House, Not really Needed. Just for remembering
##XPos 23		<-- The X Position of the House
##YPos 30		<-- The Y Position of the House
##Key 1			<-- Which key is needed to enter building 1=Default
##KName General Key	<-- Name of key which is needed, shown if player hasn't got it

*.Pls
This is the Person location Script file, it tells Green Effect where each person is and what they are called and what they say. There is a Pls file for each map file. Not all the characters in the Pls file are people some are signposts and others are lights

Eg.

1	<-- This is the number of Charcters there are in the File
##Name Taiyph	<-- The Name of the Character*
##XPos 4	<-- The X Position of the Character
##YPos 19	<-- The Y Position of the Character
##Text Hello	<-- This is What he shows when he is spoken to**
##Type 2	<-- This is the Type of Character Image***

*.Olf
This is the Objet location file, this tells Green Effect where objects are, what they are and what they do. In this version of the game they are not fully intergated

Eg.

##Name Key	<-- Name of Object
##Desc A Key	<-- Description of Object
##Extra 27001	<-- Extra Information eg. KeyVal
##X 1		<-- XPosition
##Y 1		<-- YPosition


Notes about File descriptions

* Special Names Are:
ShopKeeper		Shows the Shop Menu
SignPost		Altered Message

** Special Texts are:
Random 		Shows a random Message

*** The Different Types are:
1	Graphic 1
2	Graphic 2
3	SignPost
4	Invisible
5+	Different Light Effects
