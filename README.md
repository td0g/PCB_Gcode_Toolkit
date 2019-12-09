# td0g's PCB Gcode Tools

This program contains several tools for post-processing gcode files from Eagle.  To use, simply drag-and-drop the gcode file (with a Z-probe file if available) onto PCB_Gcode_Tools.exe and follow the command prompts.

## Tools:

1. Etch/Drill Copy

2. Etch Optimizer

3. Etch Auto-Level

4. Drill Optimizer

5. Draw Gcode Path

6. Undo Previous Tool Action

### 1. Copy

This tool simply copies the original gcode path to a new location (without deleting the original).  It is useful for cutting two identical boards on one PCB.  It can be used multiple times and will only copy the original gcode (IE. using the tool twice will result in three boards, not four).

### 2. Etch Optimizer

Eagle has a habit of creating many separate etching tool paths.  This wastes lots of time as the mill has to raise the tool, move to the next path, lower the tool, then continue.

The etch optimizer will combine all adjacent etch paths into one continuous cut, reducing travel time for the mill.  It's a simple way to increase efficiency by reducing travel time.

### 3. Etch Auto-Level

Many workpiece blanks aren't flat.  This is especially true of PCB blanks, which require precision cut depths but typically come with huge bends.  However this also becomes an issue at other times, such as engraving a poorly-planed piece of wood.

If you are able to probe the surface of the blank workpiece, then you can use this tool to apply the data you collected to a .Gcode file.  It will create a new .Gcode file which constantly adjusts the Z height of the tool to account for the unevenness of the blank.

There are two options available: Bilinear Interpolation and Bicubic Spline Interpolation.  The Bilinear Interpolation is a simpler algorithm but requires a dense probe grid.  The newer Bicubic Spline Interpolation is more fidelic to a curved PCB and provides more accuracy with less data.


### 4. Drill Optimizer

The drill order in Eagle is poorly optimized, resulting in a large and unnecessary amount of travel time.  By performing a travelling salesman algorithm on the drill order, the travel time can be greatly reduced.

Two options are available for optimizing drill order: a 'standard' algorithm and a 'genetic permutation' algorithm.  The standard algorithm simply breaks up the board into sections, sorts the drill order to go through each section in a logical path, and finally brute-forces the most efficient path through each section.  It is consistent but doesn't find the best solution.

The genetic permutation algorithm is more complicated.  It is inconsistent and not recommended for larger boards, but works very well with smaller boards.


### 5. Draw

A quick way to review the gcode and check for issues.  Black lines are etches, red lines are travel.  It can be used before and after the gcode is edited to compare the results side-by-side.

## Instructions for Etch Auto-Levelling

### Collecting Z-Probe Data

This tool does not help with collecting the Z-probe array data.  There are several methods of probing, it is up to you to implement them.

### Applying the Probe Data to the Gcode File

To apply a Z-probe array to a gcode file, drag the .CSV file with the array and the .Gcode file onto the program .exe file.  Currently, the program only accepts Z-probe data in the following format.  If you collect the data in a different format, feel free to suggest improvements to the program.

*X,Y,Z*

*0,0,0*

*0,10,0.06*

*10,0,0.08*

*10,10,0.02*
