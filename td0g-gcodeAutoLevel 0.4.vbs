'Written by Tyler Gerritsen
'vtgerritsen@gmail.com

'Run with Gcode and level map CSV
'Outputs a new Gcode with active vertical axis mapping

'The vertical position of point 0, 0 is not modified
'Please register the tool at point 0, 0 before starting gCode

'###################################################################################

			'License

'###################################################################################

'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <https://www.gnu.org/licenses/>.

'###################################################################################

			'Settings

'###################################################################################

	decimalPlaces = 2
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
	'Does not divide lines where Z is above 'ignoreAboveZ' variable
	'Bug Fix - Z correction inverted
	
'0.4
	'2019-06-24
	'Only divides lines where it is beneficial

	
'###################################################################################

			'Load Arguments (Gcode and CSV)

'###################################################################################


'Load arguments
if WScript.Arguments.Count <> 2 then
	msgbox "Please include gCode and CSV"
	WScript.quit
end if

Set Arg = WScript.Arguments

'Read files into memory
if lcase(right(arg(0),4)) = ".csv" then
	if lcase(right(arg(1),6)) = ".gcode" then
		iCSVfn = arg(0)
		Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(arg(0),1)
		csvText = Split(objFile.ReadAll,Chr(13))
		objFile.Close
		oName = replace(arg(1),".","_level.")
		Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(arg(1),1)
		gcodeText = Split(objFile.ReadAll,Chr(10))
		objFile.Close
	else
		msgbox "Please include Gcode"
		wscript.quit
	end if
elseif lcase(right(arg(1),4)) = ".csv" then
	if lcase(right(arg(0),6)) = ".gcode" then
		iCSVfn = arg(1)
		Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(arg(1),1)
		csvText = Split(objFile.ReadAll,Chr(13))
		objFile.Close
		oName = replace(arg(0),".","_level.")
		Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(arg(0),1)
		gcodeText = Split(objFile.ReadAll,Chr(10))
		objFile.Close
	else
		msgbox "Please include Gcode"
		wscript.quit
	end if
else
	msgbox "Please Include CSV"
	Wscript.quit
end if


'###################################################################################

			'Build X & Y Lists, Z Array

'###################################################################################


dim yList(1000)		'The yList will be sorted in ascending order.  y=0 is not necesarily yList(0) or yList(yListMax)
dim xList(1000)
yListSize = 0
xListSize = 0
xList(0) = ""
yList(0) = ""
yZero = ""
xZero = ""
distZero = 200
'Make xyz list
For i = LBound(csvText) to UBound(csvtext)
	c = split(csvText(i),",")
	if uBound(c) = 2 then
		if isnumeric(replace(c(0),",","")) and isnumeric(replace(c(1),",","")) and isnumeric(replace(c(2),",","")) then 
			x = CDbl(replace(c(0),",",""))
			y = CDbl(replace(c(1),",",""))
			z = CDbl(replace(c(2),",",""))
			if x = "" then x = 0
			if y = "" then y = 0
			if z = "" then z = 0
			if (x^2 + y^2)^0.5 < distZero then
				zZero = z
				distZero = (x^2 + y^2)^0.5
			end if
			
			found = false
			For j = 0 To yListSize
			  If yList(j) = y Then
				found = true
				Exit For
			  End If
			Next
			if not found then 
				yList(yListSize) = y
				if yList(yListSize) = "" then yList(yListSize) = 0
				yListSize = yListSize+1
				yList(yListSize) = ""
			end if
			found = false
			For j = 0 To xListSize
			  If xList(j) = x Then
				found = true
				Exit For
			  End If
			Next
			if not found then 
				xList(xListSize) = x
				if xList(xListSize) = "" then xList(xListSize) = 0
				xListSize = xListSize + 1
				xList(xListSize) = ""
			end if
		end if
	end if
next
xListSize = xListSize - 1
yListSize= yListsize - 1
'Sort x and y lists
for i = xListSize To 0 Step -1
    for j = 0 to i
        if xList(j)>xList(j+1) then
            temp=xList(j+1)
            xList(j+1)=xList(j)
            xList(j)=temp
			if xList(j) = "" then xList(j) = 0
        end if
    next
next 

for i = yListSize To 0 Step -1
    for j= 0 to i
        if yList(j)>yList(j+1) then
            temp=yList(j+1)
            yList(j+1)=yList(j)
            yList(j)=temp
			if yList(j) = "" then yList(j) = 0
        end if
    next
next 

'Create blank zArray
redim zArray(xListSize, yListSize)
for i = 0 to xListSize
	for j = 0 to yListSize
		zArray(i, j) = ""
	next
next

'Populate zArray
For i = LBound(csvText) to UBound(csvtext)
	c = split(csvText(i),",")
	if uBound(c) = 2 then
		if isnumeric(replace(c(0),",","")) and isnumeric(replace(c(1),",","")) and isnumeric(replace(c(2),",","")) then 
			x = CDbl(replace(c(0),",",""))
			y = CDbl(replace(c(1),",",""))
			z = CDbl(replace(c(2),",",""))
			for j = 0 to xListSize
				for k = 0 to yListSize
					if x = xList(j) and y = yList(k) then
						zArray(j, k) = round(z - zZero,decimalPlaces)
						exit for
						exit for
					end if
				next
			next
		end if
	end if
next

'Check for blank entries in zArray, exit script if any found
emptyZValues = 0
for i = 0 to xListSize
	for j = 0 to yListSize
		if zArray(i, j) = "" then emptyZValues = emptyZValues + 1
	next
next
if emptyZValues > 0 then msgbox "Missing Probe Locations Found: " & emptyZValues & vbCrLf & "Script will attempt to produce Gcode anyway" & vbCrLf & "It is recommended not to use output"

'###################################################################################

			'Write X & Y Lists, Z Array to CSV File

'###################################################################################

Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(replace(iCSVfn,".","_array."),2,true)
objFile.write ","
for i = 0 to xListSize
	objFile.write(xList(i)) & ","
next
objFile.write vbcrlf
for i = 0 to yListSize
	objFile.write yList(yListSize - i) & ","
	for j = 0 to xListSize
		objFile.Write zArray(j, yListSize - i) & ","
	next
	objFile.write vbcrlf
next
objfile.close
Set objFile = Nothing
msgbox "Array file: " & replace(iCSVfn,".","_array.")



'###################################################################################

			'Write Gcode

'###################################################################################


xIntPos = 0
yIntPos = 0
zIntPos = 0
xNewPos = 0
yNewPos = 0
zNewPos = 0
oldZFloat = 999.999
xMin = 999.999
xMax = -999.999
yMin = 999.999
yMax = -999.999
zMin = 999.999
zMax = -999.999

lastFRA = ""
lastFRB = ""

cornerCount = 0
sideCount = 0

Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(oName,2,true)

For i = LBound(gcodeText) to UBound(gcodetext)
	gc = gcodetext(i)
	if left(gc,3) <> "G00" and left(gc, 3) <> "G01" and left(gc,2) <> "G0" and left(gc, 2) <> "G1" then	'Copy these lines verbatim
		objFile.write gc ' & vbCrLf
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
		Feedrate = ""
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
					FeedrateA = " " & g(j)
				else
					FeedRateB = " " & g(j)
				end if
			end if
		next
		
		
'###################################################################################

			'Get Z Offset

'###################################################################################

'		xB 	xA 
'	yB	zBB	zAB
'	yA	zBA	zBB

		if x or y or z then
			dist = ((xNewPos-xOldPos)^2 + (yNewPos - yOldPos)^2)^0.5
			while newXFloat <> xNewPos and newYFloat <> yNewPos
			
				'1. Are we going to cross an X or Y int before reaching target?
				'2. Which line are we going to cross next?
				xLine = 0
				while newXFloat > xList(xLine)
					xLine = xLine + 1
				wend
				yLine = 0
				while (newYFloat > yList(yLine)
					yLine = yLine + 1
				wend
				if zOldPos < ignoreAboveZ or zNewPos < ignoreAboveZ then
					'We are above workpiece, just make one line
					xIntPos = xNewPos
					yIntPos = yNewPos
					zIntPos = zNewPos
				else if xNewPos > xList(xLine) or ynewPos > yList(yLine)
					'Yes, we are going to cross the line
					if (xNewPos - newXFloat) / (xList(xLine) - newXFloat) > (yNewPos - newYFloat) / (yList(yLine) - newYFloat)
						'Crossing X line first
						xIntPos = xList(xLine)
						yIntPos = (xNewpos - xOldPos) / (xIntPos - xOldPos) * (yNewPos - yOldPos) + yOldPos
						zIntPos = (zNewpos - zOldPos) / (zIntPos - zOldPos) * (zNewPos - zOldPos) + zOldPos
					else
						'Crossing Y line first
						yIntPos = yList(yLine)
						xIntPos = (yNewpos - yOldPos) / (yIntPos - yOldPos) * (xNewPos - xOldPos) + xOldPos
						zIntPos = (yNewpos - yOldPos) / (yIntPos - yOldPos) * (zNewPos - zOldPos) + zOldPos
					end if
				else
					'Nope, we are going to end this line
					xIntPos = xNewPos
					yIntPos = yNewPos
					zIntPos = zNewPos
				end if
				
				'Is point outside corner?
				if xIntPos <= xList(0) and yIntPos <= yList(0) then				
					P = zArray(0, 0)
					cornerCount = cornerCount + 1
				elseif xIntPos <= xList(0) and yIntPos >= yList(yListSize) then				
					P = zArray(0, yListSize)
					cornerCount = cornerCount + 1
				elseif xIntPos >= xList(xListSize) and yIntPos >= yList(yListSize) then				
					P = zArray(0, yListSize)
					cornerCount = cornerCount + 1
				elseif xIntPos >= xList(xListSize) and yIntPos <= yList(0) then					
					P = zArray(xListSize, 0)
					cornerCount = cornerCount + 1
					
				'Is point outside edge?
				elseif xIntPos <= xList(0) then
					sideCount = sideCount + 1
					for j = 0 to yListSize
						if yIntPos < yList(j) then
							yP = j
							exit for
						end if
					next 	
					yA = yList(yP-1)
					yB = yList(yP)
					zAB = zArray(0, yP-1)
					zBB = zArray(0, yP)
							
					P = ((yB - yIntPos)/(yB - yA))*zAB + ((yIntPos - ya)/(yb - ya))*zBB
					
				elseif xIntPos >= xList(xListSize) then
					sideCount = sideCount + 1
					for j = 0 to yListSize
						if yIntPos < yList(j) then
							yP = j
							exit for
						end if
					next 
					yA = yList(yP-1)
					yB = yList(yP)
					zAB = zArray(xListSize, yP-1)
					zBB = zArray(xListSize, yP)
					
					P = ((yB - yIntPos)/(yB - yA))*zAB + ((yIntPos - ya)/(yb - ya))*zBB
					
				elseif yIntPos <= yList(0) then
					sideCount = sideCount + 1
					for j = 0 to xListSize
						if xIntPos < xList(j) then
							xP = j
							exit for
						end if
					next
					xA = xList(xP-1)
					xB = xList(xP)
					zAA = zArray(xP-1, 0)
					zBA = zArray(xP, 0)
					
					P = ((xIntPos - xA)/(xB - xA))*zBA + ((xB - xIntPos)/(xB - xA))*zAA
					
				elseif yIntPos >= yList(ylistSize) then
					sideCount = sideCount + 1
					for j = 0 to xListSize
						if xIntPos < xList(j) then
							xP = j
							exit for
						end if
					next
					xA = xList(xP-1)
					xB = xList(xP)
					zAA = zArray(xP-1, yListSize)
					zBA = zArray(xP, yListSize)
					
					P = ((xIntPos - xA)/(xB - xA))*zBA + ((xB - xIntPos)/(xB - xA))*zAA
					
				else	'Inside Grid --> Bilinear Interpolation		
					' https://en.wikipedia.org/wiki/Bilinear_interpolation
					'http://supercomputingblog.com/graphics/coding-bilinear-interpolation/
					for j = 0 to xListSize
						if xIntPos < xList(j) then
							xP = j
							exit for
						end if
					next
					for j = 0 to yListSize
						if yIntPos < yList(j) then
							yP = j
							exit for
						end if
					next 
					xA = xList(xP-1)
					xB = xList(xP)
					yA = yList(yP-1)
					yB = yList(yP)
					zAA = zArray(xP-1, yP-1)
					zBA = zArray(xP-1, yP)
					zAB = zArray(xP, yP-1)
					zBB = zArray(xP, yP)
					
					rA = ((xIntPos - xA)/(xB - xA))*zBA + ((xB - xIntPos)/(xB - xA))*zAA
					rB = ((xIntPos - xA)/(xB - xA))*zBB + ((xB - xIntPos)/(xB - xA))*zAB
					P = ((yB - yIntPos)/(yB - yA))*rA + ((yIntPos - yA)/(yB - yA))*rB
				end if
			
		'###################################################################################

					'Resume Writing Gcode

		'###################################################################################
				
				objFile.write GType
				if x then
					newXFloat = round(xIntPos, decimalPlaces)
					objFile.Write(" X")
					objFile.Write(newXFloat)
					if newXFloat > xMax then xMax = newXFloat
					if newXFloat < xMin then xMin = newXFloat
				end if
				if y then
					newYFloat = round(yIntPos, decimalPlaces)
					objFile.Write(" Y")
					objFile.Write(newYFloat)
					if newYFloat > yMax then yMax = newYFloat
					if newYFloat < yMin then yMin = newYFloat
				end if
				if z or x or y then
					newZFloat = round(zIntPos + P,decimalPlaces)
					if newZFloat <> oldZFloat then
						objFile.Write(" Z")
						objFile.Write(newZFloat)
						oldZFloat = newZFloat
						if newZFloat > zMax then zMax = newZFloat
						if newZFloat < zMin then zMin = newZFloat
					end if
				end if
				if GType = "G0" and FeedrateA <> lastFRA then
					objFile.write FeedrateA
					lastFRA = FeedRateA
				elseif GType = "G1" and FeedrateB <> lastFRB then
					objFile.write FeedrateB 
					lastFRB = FeedRateB
				end if
				objFile.write vbCrLf
			wend
		else 
			objFile.write GType & vbCrLf
		End If
	end if
Next

'Close the file.
objFile.Close
Set objFile = Nothing
msgbox "Done!" & vbCrLf & "Levelled Gcode: " & oName & vbCrLf & "Points at or outside corners: " & cornerCount & vbCrLf & "Points at or outside edges: " & sideCount & vbNewLine & vbNewLine & "        X          Y          Z" & vbNewLine & "Max: " & xMax & "    " & yMax & "    " & zMax & vbnewline & "Min: " & xMin & "    " & yMin & "    " & zMin
