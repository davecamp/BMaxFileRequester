
Import "ModernFileRequesterNG.bmx"



' This allows you to select just a folder
Local RequestesFolder:String = RequestFolder("Select a Folder...","")



' This shows how to select 1 or more files - use shift or ctrl to select multiple files
Local files:String[] = RequestFiles("Select some file(s)","")
For Local file:String = EachIn files
	Print file
Next
