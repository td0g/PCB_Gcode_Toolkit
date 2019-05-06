# Gcode Auto-Level

## Description

Many workpiece blanks aren't flat.  This is especially true of PCB blanks, which require precision cut depths but typically come with huge bends.  However this also becomes an issue at other times, such as engraving a poorly-planed piece of wood.

If you are able to probe the surface of the blank workpiece, then you can use this tool to apply the data you collected to a .Gcode file.  It will change the Z height of the tool to account for the unevenness of the blank.

For more information, see www.td0g.ca/projects/cnc_mini_mill#autolevel

## Instructions

### Collecting Z-Probe Data

This tool does not help with collecting the Z-probe array data.  There are several methods of probing, it is up to you to implement them.

### Convert to Usable CSV File

The gcodeAutoLevel expects a .CSV file with three columns arranged in the following order: X, Y, Z.

If you have a text file with a list of XYZ coordinates, you can use the convertTextToCSV.vbs script to generate a formatted CSV file.  Just drag the text file with the Z-probe array data onto the convertTextToCSV.vbs script

### Applying the Probe Data to the Gcode File

To apply a Z-probe array to a gcode file, drag the .CSV file with the array and the .Gcode file onto the gcodeAutoLevel.vbs script.

## Settings

To adjust the script settings, open the gcodeAutoLevel.vbs script in a text editor.  The following settings can be found near the top:

- decimalPlaces   *(Precision of output file - Default is 2 decimal places)*
- maxLineLength   *(Maximum length of a line to leave undivided - Default is 3mm)*
- lineLength      *(If line is going to be divided, maximum length of line segments - Default is 2mm)*

## Example

### 1.

If you have a text file that contains data like this, then drag the file onto the convertTextToCSV.vbs script:

*X:0 Y:0 Z:0*

*X:0 Y:10 Z:0.06*

*X:10 Y:0 Z:0.08*

*X:10 Y:10 Z:0.02*

It will generate a .CSV that is formatted as shown below.  If you already have a .CSV file that looks like this, then proceed to step 2:

*X,Y,Z*

*0,0,0*

*0,10,0.06*

*10,0,0.08*

*10,10,0.02*

### 2.

Select the .CSV file AND the .gcode file.  Drag them onto the gcodeAutoLevel.csv script.  It will display a couple notifications before completion and one when it is completed.  If your .gcode file has thousands of lines, then be patient while it works - it may take several minutes.

The script also generates another .CSV file which arranges the probe data into a matrix.  Feel free to delete when you are finished.
