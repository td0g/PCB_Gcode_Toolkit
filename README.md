# td0g's PCB Gcode Toolkit

This program contains several tools for post-processing gcode files from Eagle.  To use, simply drag-and-drop the gcode file (along with a text file containing Z-probe data if available) onto PCB_Gcode_Tools.exe and follow the command prompts.

https://youtu.be/JNCeBOY9sTc

I began developing this program back when there weren't really any free solutions available.  It is simple, runs from a command line interface and requires little user input.  It also uses clever algorithms to improve efficiency: The **bicubic spline interpolation** means it needs fewer probe points to auto-level a pcb board, the **divide-and-conquer TSP algorithm** is fast and effective, and the **etch optimizer** eliminates almost all unnecessary travel movements.

This tool has currently not been tested on KiCad gcode.

## Tools:

1. Etch/Drill Copy

2. Etch Optimizer

3. Etch Auto-Level

4. Drill Optimizer

5. Backlash Compensation

6. Draw Gcode Path

7. Undo Previous Tool Action

### 1. Copy

This tool simply copies the original gcode path to a new location (without deleting the original).  It is useful for cutting two identical boards on one PCB.  It can be used multiple times and will only copy the original gcode (IE. using the tool twice will result in three boards, not four).

### 2. Etch Optimizer

Eagle has a habit of creating many separate etching tool paths.  This wastes lots of time as the mill has to raise the tool, move to the next path, lower the tool, then continue.

The etch optimizer will combine all adjacent etch paths into one continuous cut, reducing travel time for the mill.  It's a simple way to increase efficiency by reducing travel time.

### 3. Etch Auto-Level

Many workpiece blanks aren't flat.  This is especially true of PCB blanks, which require precision cut depths but typically come with huge bends.  However this also becomes an issue at other times, such as engraving a poorly-planed piece of wood.

If you are able to probe the surface of the blank workpiece, then you can use this tool to apply the data you collected to a .Gcode file.  It will create a new .Gcode file which constantly adjusts the Z height of the tool to account for the unevenness of the blank.

There are two options available: Bilinear Interpolation and Bicubic Spline Interpolation.  Descriptions of both algorithms are available in PDF format.  The Bilinear Interpolation is a simpler algorithm but requires a dense probe grid. 

![Bilinear Interpolation - Courtesy of Wikipedia](https://upload.wikimedia.org/wikipedia/commons/thumb/d/dd/Interpolation-bilinear.svg/220px-Interpolation-bilinear.svg.png)

The newer Bicubic Spline Interpolation is more fidelic to a curved PCB and provides more accuracy with less data.  Not only is the interpolation algorithm better, it produces more efficient gcode (fewer added commands) than Bilinear Interpolation.

![Bicubic Spline Interpolation - Courtesy of Wikipedia](https://upload.wikimedia.org/wikipedia/commons/thumb/f/f5/Interpolation-bicubic.svg/220px-Interpolation-bicubic.svg.png)


### 4. Drill Optimizer

The drill order in Eagle is poorly optimized, resulting in a large and unnecessary amount of travel time.  By performing a travelling salesman algorithm on the drill order, the travel time can be greatly reduced.

Two options are available for optimizing drill order: a **genetic permutation** algorithm and a **divide-and-conquer** algorithm.  If in doubt, try both on a board.  The program will not make any changes if it can't find a better solution.

The genetic permutation algorithm processes the entire board at once.  It is inconsistent and may fail on larger boards, but will generally provide a very good solution.  The divide-and-conquer algorithm simply breaks up the board into sections, selects an appropriate algorithm for each section (genetic permutation or brute-force), and finally performs each section in a logical order.


### 5. Backlash Compensation

Currently this feature only affects Z-axis backlash.  When the Z-axis is moved downward, the backlash compensation function will 'overshoot' the Z-position, then return to the normal Z-position.  This forces the tool down to the correct position. 

### 6. Draw

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

## Feedback

If you have any question, suggestions, or bug finds, please feel free to contact me.  I am continually developing new features and improvements, and am happy to hear from the community.  Thanks!


## License

Software is licensed under a [GNU GPL v3 License](https://www.gnu.org/licenses/gpl-3.0.txt)
