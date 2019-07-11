'Written by Tyler Gerritsen
'vtgerritsen@gmail.com

'Run with Gcode and level map CSV
'Outputs a new Gcode with active vertical axis mapping

'The vertical position of point 0, 0 is not modified
'Please register the tool at point 0, 0 before starting gCode

'###################################################################################

			'Settings

'###################################################################################

	decimalPlaces = 2
	maxLineLength = 3
	lineLength = 2
	ignoreAboveZ = 1

'###################################################################################

			'Changelog

'###################################################################################


'0.1
	'2017-02-10
	'Functional
	'Does not work with Relative Movements
'0.2
	'2019-05-19
	'Does not repeat same Z position
	'Summarizes dimensions of mill volume at end
'0.3
	'2019-05-25
	'Does not divide lines where Z is above set level

	
'###################################################################################

			'Load Arguments (Gcode)

'###################################################################################


'Load arguments
if WScript.Arguments.Count <> 1 then
	msgbox "Please include gCode"
	WScript.quit
end if

Set Arg = WScript.Arguments

'Read files into memory
if lcase(right(arg(0),6)) = ".gcode" then
	Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(arg(0),1)
	gcodeText = Split(objFile.ReadAll,Chr(10))
	objFile.Close
else
	msgbox "Please include Gcode"
	wscript.quit
end if


'###################################################################################

			'Write Gcode

'###################################################################################

xOldPos = 0.0
yOldPos = 0.0
zOldPos = 0.0
xNewPos = 0.0
yNewPos = 0.0
zNewPos = 0.0
speedZero = 60.0
speedOne = 60.0
zeroOne = 1
xMin = 999.999
xMax = -999.999
yMin = 999.999
yMax = -999.999
zMin = 999.999
zMax = -999.999
timeTotal = 0.0
distTotal = 0.0
distThis = 0.0




For i = LBound(gcodeText) to UBound(gcodetext)
	gc = gcodetext(i)
	if left(gc,3) <> "G00" and left(gc, 3) <> "G01" and left(gc,2) <> "G0" and left(gc, 2) <> "G1" then	'Copy these lines verbatim
		timeTotal = timeTotal
	elseif left(gc,1) = "G" then
		gc = replace(gc, "G00", "G0")
		gc = replace(gc, "G01", "G1")
		xOldPos = xNewPos
		yOldPos = yNewPos
		zOldPos = zNewPos
		gc = replace(gc,Chr(10),"")	'Remove newline chars
		gc = replace(gc,Chr(13),"")	'Remove newline chars
		g = split(gc," ")
		x = 0
		y = 0
		z = 0
		GType = ""
		for j = LBound(g) to UBound(g)
			g(j) = replace(g(j)," ","")
			if inStr(g(j),"G") > 0 then 
				GType = g(j)
			elseif inStr(g(j),"X") > 0 then 
				xNewPos = CDbl(replace(g(j),"X",""))
				x = 1
			elseif inStr(g(j),"Y") > 0 then 
				yNewPos = CDbl(replace(g(j),"Y",""))
				y = 1
			elseif inStr(g(j),"Z") > 0 then 
				zNewPos = CDbl(replace(g(j),"Z",""))
				z = 1
			elseif inStr(g(j),"F") > 0 then 
				if GType = "G0" then
					speedZero = CDbl(replace(g(j),"F",""))
					zeroOne = 0
				else
					speedOne = CDbl(replace(g(j),"F",""))
					zeroOne = 1
				end if
			end if
		next
		if xMax < xNewPos then xMax = xNewPos
		if xMin > xNewPos then xMin = xNewPos
		if yMax < yNewPos then yMax = yNewPos
		if yMin > yNewPos then yMin = yNewPos
		if zMax < zNewPos then zMax = zNewPos
		if zMin > zNewPos then zMin = zNewPos
		distThis = ((xNewPos - xOldPos)^2 + (yNewpos - yOldPos) ^ 2 + (zNewPos - zOldPos) ^ 2) ^ 0.5
		distTotal = distTotal + distThis
		if zeroOne = 0 then
			timeTotal = timeTotal + distThis / speedZero * 60
		else
			timeTotal = timeTotal + distThis / speedOne * 60
		end if
	end if
Next

'Close the file.
objFile.Close
Set objFile = Nothing

msgbox "Done!" & vbCrLf & vbNewLine & "        X          Y          Z" & vbNewLine & "Max: " & xMax & "    " & yMax & "    " & zMax & vbnewline & "Min: " & xMin & "    " & yMin & "    " & zMin & vbNewline & vbNewline & "Total Distance: " & distTotal & vbNewline & "Total Time: " & timeTotal
