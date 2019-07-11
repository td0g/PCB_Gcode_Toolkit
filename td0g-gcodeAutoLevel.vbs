'Written by Tyler Gerritsen
'vtgerritsen@gmail.com

'Run with Gcode and level map CSV
'Outputs a new Gcode with active vertical axis mapping

'The vertical position of point 0, 0 is not modified
'Please register the tool at point 0, 0 before executing gCode

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
	'Now accepts .csv OR .txt files
	'Bilinear Interpolation Bug Fix

	
'###################################################################################

			'Load Arguments (Gcode and CSV)

'###################################################################################


'Load arguments
if WScript.Arguments.Count <> 2 then
	msgbox "Please include gCode and CSV" & vbNewLine & vbNewLine & "For more information, please visit github.com/td0g/gcode_auto_level"
	WScript.quit
end if

Set Arg = WScript.Arguments

'Read files into memory
if lcase(right(arg(0),4)) = ".csv" or lcase(right(arg(0),4)) = ".txt" then
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
		msgbox "Please include Gcode" & vbNewLine & vbNewLine & "For more information, please visit github.com/td0g/gcode_auto_level"
		wscript.quit
	end if
elseif lcase(right(arg(1),4)) = ".csv" or lcase(right(arg(1),4)) = ".txt" then
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
		msgbox "Please include Gcode" & vbNewLine & vbNewLine & "For more information, please visit github.com/td0g/gcode_auto_level"
		wscript.quit
	end if
else
	msgbox "Please Include CSV" & vbNewLine & vbNewLine & "For more information, please visit github.com/td0g/gcode_auto_level"
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
if emptyZValues > 0 then msgbox "Missing Probe Locations Found: " & emptyZValues & vbCrLf &_ 
	"Script will attempt to produce Gcode anyway" & vbCrLf & "It is recommended not to use output"

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




'###################################################################################

			'Process Gcode

'###################################################################################


xIntPos = 0.0
yIntPos = 0.0
zIntPos = 0.0
xLastIntPos = 0.0
yLastIntPos = 0.0
zLastIntPos = 0.0
xTarget = 0.0
yTarget = 0.0
zTarget = 0.0
xLastTarget = 0.0
yLastTarget = 0.0
zLastTarget = 0.0
zRoundedFloatOld = 999.999
xMin = 999.999
xMax = -999.999
yMin = 999.999
yMax = -999.999
zMin = 999.999
zMax = -999.999
firstMove = true

lastFRA = ""
lastFRB = ""

newTotalCount = 0
oldTotalCount = 0
cornerCount = 0
edgeCount = 0

oName = replace(oName, ".gcode", ".tmp")
Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(oName,2,true)

'Get starting Z position
For i = LBound(gcodeText) to UBound(gcodetext)
	gc = gcodetext(i)
	if left(gc,3) = "G00" or left(gc,3) = "G01" or left(gc,3) = "G0 " or left(gc,3) = "G1 " then
		g = split(gc," ")
		for j = LBound(g) to UBound(g)
			g(j) = replace(g(j)," ","")
			if inStr(g(j),"Z") > 0 then 
				zLastTarget = CDbl(replace(g(j),"Z",""))
				zTarget = zLastTarget
				exit for
				exit for
			end if
		next
	end if
next
objFile.Close
Set objFile = Nothing


Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(oName,2,true)

For i = LBound(gcodeText) to UBound(gcodetext)
	gc = gcodetext(i)
	if left(gc,3) <> "G00" and left(gc, 3) <> "G01" and left(gc,3) <> "G0 " and left(gc, 3) <> "G1 " then	'Copy these lines verbatim
		objFile.write gc
	elseif left(gc,1) = "G" then
		oldTotalCount = oldTotalCount + 1
		gc = replace(gc, "G00", "G0")
		gc = replace(gc, "G01", "G1")
		gc = replace(gc,Chr(10),"")	'Remove newline chars
		gc = replace(gc,Chr(13),"")	'Remove newline chars
		g = split(gc," ")
		x = 0
		y = 0
		z = 0
		Feedrate = ""
		GType = ""
		xLastTarget = xTarget
		yLastTarget = yTarget
		zLastTarget = zTarget
		for j = LBound(g) to UBound(g)
			g(j) = replace(g(j)," ","")
			if inStr(g(j),"G") > 0 then 
				GType = g(j)
			elseif inStr(g(j),"X") > 0 then 
				xTarget = CDbl(replace(g(j),"X",""))
				if xTarget <> xLastTarget or firstMove then x = 1
			elseif inStr(g(j),"Y") > 0 then 
				yTarget = CDbl(replace(g(j),"Y",""))
				if yTarget <> yLastTarget or firstMove then y = 1
			elseif inStr(g(j),"Z") > 0 then 
				zTarget = CDbl(replace(g(j),"Z",""))
				if zTarget <> zLastTarget or firstMove then z = 1
			elseif inStr(g(j),"F") > 0 then 
				if GType = "G0" then
					FeedrateA = " " & g(j)
				else
					FeedRateB = " " & g(j)
				end if
			end if
		next
		
		
'###################################################################################

			'Calculate position of intermediate point

'###################################################################################

		if x or y or z or firstMove then
			dist = ((xTarget-xLastTarget)^2 + (yTarget - yLastTarget)^2 + (zTarget - zLastTarget)^2)^0.5
			while xIntPos <> xTarget or yIntPos <> yTarget or zIntPos <> zTarget or firstMove
				xLastIntPos = xIntPos
				yLastIntPos = yIntPos
				zLastIntPos = zIntPos
'Only moving Z - go straight there
				if x = 0 and y = 0 and firstMove = false then
					zIntPos = zTarget
'We are above workpiece or this is the first move, just make one line
				elseif zLastTarget > ignoreAboveZ or zTarget > ignoreAboveZ or firstMove then
					firstMove = false
					xIntPos = xTarget
					yIntPos = yTarget
					zIntPos = zTarget
'1. Are we going to cross an X or Y int before reaching target?
'2. Which line are we going to cross first?
				else
					if x then
						if xLastIntPos < xList(0) or xTarget =< xList(0) then
							if xLastIntPos =< xList(0) and xTarget =< xList(0) then
								xIntPos = xTarget
							else
								xIntPos = xList(0)
							end if
						elseif xLastIntPos > xList(xlistsize) or xTarget => xList(xlistsize) then
							if xLastIntPos => xList(xlistsize) and xTarget => xList(xlistsize) then
								xIntPos = xTarget
							else
								xIntPos = xList(xlistsize)
							end if
						else 
							if xTarget > xLastTarget then
								xLine = 0
								while xIntPos => xList(xLine)
									xLine = xLine + 1
								wend
							else								
								xLine = xListSize
								while xIntPos =< xList(xLine)
									xLine = xLine - 1
								wend
							end if

							if xTarget > xIntPos then
								if xTarget > xList(xLine) then 
									xIntPos = xList(xLine)
								else
									xIntPos = xTarget
								end if
							elseif xTarget < xIntPos then
								if xTarget < xList(xLine) then 
									xIntPos = xList(xLine)
								else
									xIntPos = xTarget
								end if
							end if
						end if
					else
						xIntPos = xTarget 'X is not moving, just set to final target
					end if
					
					if y then
						if yLastIntPos < yList(0) or yTarget =< yList(0) then
							if yLastIntPos =< yList(0) and yTarget =< yList(0) then
								yIntPos = yTarget
							else
								yIntPos = yList(0)
							end if
						elseif yLastIntPos > yList(ylistsize) or yTarget => yList(ylistsize) then
							if yLastIntPos => yList(ylistsize) and yTarget => yList(ylistsize) then
								yIntPos = yTarget
							else
								yIntPos = yList(ylistsize)
							end if
						else 
							if yTarget > yLastTarget then
								yLine = 0
								while yIntPos => yList(yLine)
									yLine = yLine + 1
								wend
							else								
								yLine = yListSize
								while yIntPos =< yList(yLine)
									yLine = yLine - 1
								wend
							end if
							if yTarget > yIntPos then
								if yTarget > yList(yLine) then 
									yIntPos = yList(yLine)
								else
									yIntPos = yTarget
								end if
							elseif yTarget < yIntPos then
								if yTarget < yList(yLine) then 
									yIntPos = yList(yLine)
								else
									yIntPos = yTarget
								end if
							end if
						end if
					else
						yIntPos = yTarget 'Y is not moving, just set to final target
					end if
					
'What intPos comes first?
					if y = 0 then	'Y is not moving, X comes first
						yIntPos = yLastTarget
						zIntPos = (xIntPos - xLastTarget) / (xTarget - xLastTarget) * (zTarget - zLastTarget) + zLastTarget
					elseif x = 0 then	'X is not moving, Y comes first
						xIntPos = xLastTarget
						zIntPos = (yIntPos - xLastTarget) / (yTarget - yLastTarget) * (zTarget - zLastTarget) + zLastTarget
					elseif (xIntPos - xLastIntPos) / (xTarget - xLastTarget) > (yIntPos - yLastIntPos) / (yTarget - yLastTarget) then 'Crossing Y line first
						xIntPos = (yIntPos - yLastTarget) / (yTarget - yLastTarget) * (xTarget - xLastTarget) + xLastTarget
						zIntPos = (yIntPos - xLastTarget) / (yTarget - yLastTarget) * (zTarget - zLastTarget) + zLastTarget
					else 'Crossing X line first
						yIntPos = (xIntPos - xLastTarget) / (xTarget - xLastTarget) * (yTarget - yLastTarget) + yLastTarget
						zIntPos = (xIntPos - xLastTarget) / (xTarget - xLastTarget) * (zTarget - zLastTarget) + zLastTarget
					end if
				end if
				
'###################################################################################

			'Apply Bilinear Interpolation

'###################################################################################

'Is this a hop?
				if zIntPos > ignoreAboveZ then
					P = 0
'Is point outside corner?
				elseif xIntPos <= xList(0) and yIntPos <= yList(0) then				
					P = zArray(0, 0)
					cornerCount = cornerCount + 1
				elseif xIntPos <= xList(0) and yIntPos >= yList(yListSize) then				
					P = zArray(0, yListSize)
					cornerCount = cornerCount + 1
				elseif xIntPos >= xList(xListSize) and yIntPos >= yList(yListSize) then				
					P = zArray(xListSize, yListSize)
					cornerCount = cornerCount + 1
				elseif xIntPos >= xList(xListSize) and yIntPos <= yList(0) then					
					P = zArray(xListSize, 0)
					cornerCount = cornerCount + 1
					
'Is point outside edge?
				elseif xIntPos <= xList(0) then
					edgeCount = edgeCount + 1
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
					edgeCount = edgeCount + 1
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
					edgeCount = edgeCount + 1
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
					edgeCount = edgeCount + 1
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
					
'Inside Grid --> Bilinear Interpolation		
' https://en.wikipedia.org/wiki/Bilinear_interpolation
'http://supercomputingblog.com/graphics/coding-bilinear-interpolation/
				else	
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

			'Write to Gcode File

'###################################################################################
				
				zRoundedFloat = round(zIntPos + P,decimalPlaces)
				if x or y or zRoundedFloat <> zRoundedFloatOld then 
					objFile.write GType
					newTotalCount = newTotalCount + 1
				end if
				if x then
					xRoundedFloat = round(xIntPos, decimalPlaces)
					objFile.Write(" X")
					objFile.Write(xRoundedFloat)
					if xRoundedFloat > xMax then xMax = xRoundedFloat
					if xRoundedFloat < xMin then xMin = xRoundedFloat
				end if
				if y then
					yRoundedFloat = round(yIntPos, decimalPlaces)
					objFile.Write(" Y")
					objFile.Write(yRoundedFloat)
					if yRoundedFloat > yMax then yMax = yRoundedFloat
					if yRoundedFloat < yMin then yMin = yRoundedFloat
				end if
				if z or x or y then
					zRoundedFloat = round(zIntPos + P,decimalPlaces)
					if zRoundedFloat <> zRoundedFloatOld then
						objFile.Write(" Z")
						objFile.Write(zRoundedFloat)
						if zRoundedFloat > zMax then zMax = zRoundedFloat
						if zRoundedFloat < zMin then zMin = zRoundedFloat
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
				zRoundedFloatOld = zRoundedFloat
			wend
		'else 
			'objFile.write GType & vbCrLf
		End If
	end if
Next

'###################################################################################

			'Wrap it up

'###################################################################################

'Close the file
objFile.Close
Set objFile = Nothing

'Rename .tmp to .gcode
Dim Fso
Set Fso = WScript.CreateObject("Scripting.FileSystemObject")
if Fso.FileExists(replace(oName,".tmp",".gcode")) then Fso.DeleteFile replace(oName,".tmp",".gcode")
Fso.MoveFile oName, replace(oName,".tmp",".gcode")

'Notify user
percentInside = 0.00
percentInside = 100 * (1 - (edgeCount + cornerCount) / oldTotalCount)
msgbox "Done!" & vbCrLf & vbCrLf & "Levelled Gcode: " & replace(oName,".tmp",".gcode") & vbCrLf & vbCrLf & "Array file: " & replace(iCSVfn,".","_array.") &_ 
	vbCrLf & vbCrLf & "Points before / after:  " & oldTotalCount & " / " & newTotalCount &_
	vbCrLf & "Points at or outside corners:  " & cornerCount & vbCrLf & "Points at or outside edges:    " & edgeCount &_
	vbCrLf & "Percentage inside probed area:  " & round(percentInside,1) & "%" &_
	vbNewLine & vbNewLine & "           X           Y           Z" & vbNewLine & "Max:   " & left(xMax & "            ",12) & left(yMax & "            ",12) & zMax & vbnewline & "Min:   " &_ 
	left(xMin & "            ",12) & left(yMin & "            ",12) & zMin
