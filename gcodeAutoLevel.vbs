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

'###################################################################################

			'Changelog

'###################################################################################


'0.1
	'2017-02-10
	'Functional
	'Does not work with Relative Movements

	
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

lastFRA = ""
lastFRB = ""

cornerCount = 0
sideCount = 0

Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(oName,2,true)
msgbox "Levelled Gcode: " & oName
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
			'Get line length
			dist = ((xNewPos-xOldPos)^2 + (yNewPos - yOldPos)^2)^0.5
			segs = 1
			segDist = dist
			if dist > maxLineLength then	'If line length is greater than maxLineLength, divide until it is smaller than lineLength
				while segDist > lineLength
					segs = segs + 1
					segDist = dist / segs
				wend
			end if
			for k = 1 to segs
				'Get intermediate position (even if line is not divided into multiple segments)
				xIntPos = xOldPos + (xNewPos - xOldpos) / segs * k
				yIntPos = yOldPos + (yNewPos - yOldpos) / segs * k
				zIntPos = zOldPos + (zNewPos - zOldpos) / segs * k
				
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
					objFile.Write(" X")
					objFile.Write(round(xIntPos, decimalPlaces))
				end if
				if y then
					objFile.Write(" Y")
					objFile.Write(round(yIntPos, decimalPlaces))
				end if
				if z or x or y then
					objFile.Write(" Z")
					objFile.Write(round(zIntPos - P,decimalPlaces))
				end if
				if GType = "G0" and FeedrateA <> lastFRA then
					objFile.write FeedrateA
					lastFRA = FeedRateA
				elseif GType = "G1" and FeedrateB <> lastFRB then
					objFile.write FeedrateB 
					lastFRB = FeedRateB
				end if
				objFile.write vbCrLf
			next
		else 
			objFile.write GType & vbCrLf
		End If
	end if
Next

'Close the file.
objFile.Close
Set objFile = Nothing

msgbox "Done!" & vbCrLf & "Points at or outside corners: " & cornerCount & vbCrLf & "Points at or outside edges: " & sideCount
