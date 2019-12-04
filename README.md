# td0g's PCB Milling Tools

This program contains several tools for post-processing gcode files from Eagle.  To use, simply drag-and-drop the gcode file onto PCB_Gcode_Tools.exe and follow the command prompts.

## Tools:

1. Etch Auto-Level

2. Etch Optimizer

3. Drill Optimizer

4. Etch/Drill Copy

5. Draw Gcode Path

### 1. Etch Auto-Level

Many workpiece blanks aren't flat.  This is especially true of PCB blanks, which require precision cut depths but typically come with huge bends.  However this also becomes an issue at other times, such as engraving a poorly-planed piece of wood.

If you are able to probe the surface of the blank workpiece, then you can use this tool to apply the data you collected to a .Gcode file.  It will create a new .Gcode file which constantly adjusts the Z height of the tool to account for the unevenness of the blank.

### 2. Etch Optimizer

Eagle has a habit of creating many separate etching tool paths.  This wastes lots of time as the mill has to raise the tool, move to the next path, lower the tool, then continue.

The etch optimizer will combine all adjacent etch paths into one continuous cut, reducing travel time for the mill.  It's a simple way to increase efficiency.

### 3. Drill Optimizer

The drill order in Eagle is poorly optimized, resulting in a large and unnecessary amount of travel time.  By performing a travelling salesman algorithm on the drill order, the travel time can be reduced.

Two options are available for optimizing drill order: a 'standard' algorithm and a 'genetic permutation' algorithm.  The standard algorithm simply breaks up the board into sections, sorts the drill order to go through each section in a logical path, and finally brute-forces the most efficient path through each section.  It is consistent but doesn't find the best solution.

The genetic permutation algorithm is more complicated.  It is inconsistent and not recommended for larger boards, but works very well with smaller boards.

### 4. Copy

This tool simply copies the original gcode path to a new location (without deleting the original).  It is useful for cutting two identical boards on one PCB.

### 5. Draw

A quick way to review the gcode and check for issues.  Black lines are etches, red lines are travel.

## Instructions for Etch Auto-Levelling

### Collecting Z-Probe Data

This tool does not help with collecting the Z-probe array data.  There are several methods of probing, it is up to you to implement them.

### Convert to Usable CSV File

The gcodeAutoLevel expects a .CSV file with three columns arranged in the following order: X, Y, Z.

If you have a text file with a list of XYZ coordinates, you can use the convertTextToCSV.vbs script to generate a formatted CSV file.  Just drag the text file with the Z-probe array data onto the convertTextToCSV.vbs script

### Applying the Probe Data to the Gcode File

To apply a Z-probe array to a gcode file, drag the .CSV file with the array and the .Gcode file onto the gcodeAutoLevel.vbs script.

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
