'Written by Tyler Gerritsen
'vtgerritsen@gmail.com

'Drag a text file containing X, Y, and Z coordinates from a Z-probe grid,
'Output to a .CSV file compatible with the gcode_auto_level script

'0.1 2017-02-02

	
Set Arg = WScript.Arguments
	
	Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(arg(0),1)
    strFileText = objFile.ReadAll
    objFile.Close
	
        'Split the file at the new line character. *Use the Line Feed character (Char(10))
    arrFileText = Split(strFileText,Chr(10))
        'Open the file for writing.
	oFile =  replace(arg(0),right(arg(0),3), "csv")
    Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(oFile,2,true)
        'Loop through the array of lines looking for lines to keep.
    For i = LBound(arrFileText) to UBound(arrFileText)
		if len(arrFileText(i)) > 15 then
			arrFileText(i) = Right(arrFileText(i), len(arrFileText(i)) - 15)
			If left(arrFileText(i), 2) = "X:" Then
				arrFileText(i) = Replace(arrFileText(i),"Y:", ",")
				arrFileText(i) = Replace(arrFileText(i),"Z:", ",")
				arrFileText(i) = Replace(arrFileText(i),"X:", "")
				arrFileText(i) = Replace(arrFileText(i)," ", "")
				objFile.Write(arrFileText(i))
			End If
		End If
    Next
        'Close the file.
    objFile.Close
    Set objFile = Nothing