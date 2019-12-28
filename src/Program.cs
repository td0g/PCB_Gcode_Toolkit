/*
Written by Tyler Gerritsen
vtgerritsen@gmail.com

Run with Gcode and level map CSV
Outputs a new Gcode with active vertical axis mapping

The vertical position of point 0, 0 is not modified
Please register the tool at point 0, 0 before starting gCode
0.1 2017-02-10
	Functional
	Does not work with Relative Movements
0.2 2019-05-19
	Does not repeat same Z position
	Summarizes dimensions of mill volume at end
0.3 2019-05-25
	Does not divide lines where Z is above set level
0.4 2019-06-01
    Bug Fixes
0.5 2019-07-05
    Ported to C#
0.6 2019-08-30
    Fixed bugs, added total distance tool
0.7 2019-09-16
    Fixed bugs
    Added etch joining algorithm (improved etching speed).
    Added GA Algorithm Option
0.7.1 2019-09-22
    Fixed etch optimizer bugs (didn't complete loop, feedrate issues)
    Can now optimize without level gcode
0.7.2 2019-10-19
    Fixed hole optimizer bug (Lost one hole every time)
    No renaming of filenames
0.7.3 2019-10-24
    LEVELLING: Changed to use basic interpolation when crossing a line, bilinear interpolation was not working well
    LEVELLING: Fixed levelling array bug
0.7.4   2019-11-05
    New COPY feature
0.7.5   2019-11-09
    Simplified menu structure
0.8     2019-11-20
    New Bicubic Spline Interpolation
0.9     2019-12-03
    td0g hole-drill algorithm now uses genetic algorithm for densely-populated sectors
    Genetic algorithm doesn't overwrite if the original gcode was faster
    General code cleanup
0.10    2019-12-03
    Can DRAW gcode path
    Speed improvement to etch optimization algorithm
0.11    2019-12-07
    New UNDO feature
    Retains hole drill depth from original gcode
    Fixed Bilinear Interpolation bug
    Added Bilinear Interpolation heatmaps
0.12    2019-12-27
    New BACKLASH COMPENSATION feature
    Added heatmap to DRAW function
    Improvments to Bicubic Spline Interpolation algorithm
    Program version printed in output gcode
    Minor DRAW bug fix (zMin == zMax)
    Minor Bilinear Interpolation bug fix
    Bicubic Spline EXTRAPOLATION
    ETCH OPTIMIZATION revamp (much less memory usage)
    New OPTIONS menu


*/

using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

class Program
{
    public static string currentVersion = "0.12";
    public static void Main(string[] args)
    {
        //########################################################################################
        //                          Parse Args
        //########################################################################################


        Console.WriteLine("PCB Gcode Tools " + currentVersion);
        Console.WriteLine("Written by Tyler Gerritsen\nwww.td0g.ca\n");
        Console.ForegroundColor = ConsoleColor.White;
        string gcodePath = "";
        string levelPath = "";
        for (int i = 0; i < args.Length; i++)
        {
            if (("    "+args[i]).Substring(args[i].Length).ToLower() == ".csv" || ("    "+args[i]).Substring(args[i].Length).ToLower() == ".txt") levelPath = args[i];
            else if (("      "+args[i]).Substring(args[i].Length).ToLower() == ".gcode") gcodePath = args[i];
        }
        Console.Write("Parameters:\n  Gcode File (string): ");
        if (gcodePath == "") Console.Write("  NOT FOUND");
        else Console.Write(gcodePath);
        Console.Write("\n  Probe Data File (string): ");
        if (levelPath == "") Console.Write("  NOT FOUND");
        else Console.Write(levelPath);
        if (gcodePath == "")
        {
            Console.WriteLine("\nPress any key to continue...");
            Console.ReadKey();
            Environment.Exit(-1);
        }
        Console.WriteLine("\n");

        //########################################################################################
        //                          Run
        //########################################################################################

        char key = 'a';
        double originalLength = gcodeEdit.showStats(gcodePath);
        printHelp();
        while (key >= 'A' && key <= 'z')
        {
            key = Console.ReadKey().KeyChar;
            
            if (key == 'c')    //Copy gcode
            {
                Console.WriteLine();
                Console.Write("x Offset: ");
                double x = double.Parse(Console.ReadLine());
                Console.Write("y Offset: ");
                double y = double.Parse(Console.ReadLine());
                Console.ForegroundColor = ConsoleColor.Red;
                if (!gcodeEdit.saveUndoPoint(gcodePath)) Console.WriteLine("UNABLE TO CREATE SAVE POINT!\nContinuing without saving");
                Console.ForegroundColor = ConsoleColor.White;
                gcodeEdit.copyGcode(gcodePath, x, y);
                originalLength = gcodeEdit.showStats(gcodePath);
                printHelp();
            }
            /*
            else if (key == 'e') //Optimize etch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                if (!gcodeEdit.saveUndoPoint(gcodePath)) Console.WriteLine("UNABLE TO CREATE SAVE POINT!\nContinuing without saving");
                Console.ForegroundColor = ConsoleColor.White;
                while (gcodeEdit.combineEtches(gcodePath, 0.01)) { };
                while (gcodeEdit.combineEtches(gcodePath, 0.05)) { };
                while (gcodeEdit.combineEtches(gcodePath, 0.1)) { };
                while (gcodeEdit.combineEtches(gcodePath, 0.2)) { };
                originalLength = gcodeEdit.showStats(gcodePath);
                printHelp();
            }
            */
            else if (key == 'e') //Optimize etch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                if (!gcodeEdit.saveUndoPoint(gcodePath)) Console.WriteLine("UNABLE TO CREATE SAVE POINT!\nContinuing without saving");
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine();
                gcodeEdit.combineEtches(gcodePath, 0.05, false);
                Console.WriteLine("33%");
                gcodeEdit.combineEtches(gcodePath, 0.1, false);
                Console.WriteLine("66%");
                gcodeEdit.combineEtches(gcodePath, 0.2, true);
                originalLength = gcodeEdit.showStats(gcodePath);
                printHelp();
            }
            else if (key == 'k')    //Level with bilinear interpolation
            {
                if (levelPath != "")
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    if (!gcodeEdit.saveUndoPoint(gcodePath)) Console.WriteLine("UNABLE TO CREATE SAVE POINT!\nContinuing without saving");
                    Console.ForegroundColor = ConsoleColor.White;
                    gcodeEdit.levelGcode(gcodePath, levelPath, false);
                }
                else
                {
                    Console.WriteLine("No levelling data found!");
                    Console.WriteLine("Please include levelling data with gcode");
                }
                originalLength = gcodeEdit.showStats(gcodePath);
                printHelp();
            }
            else if (key == 'l')    //Level with bicubic spline interpolation
            {
                if (levelPath != "")
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    if (!gcodeEdit.saveUndoPoint(gcodePath)) Console.WriteLine("UNABLE TO CREATE SAVE POINT!\nContinuing without saving");
                    Console.ForegroundColor = ConsoleColor.White;
                    gcodeEdit.levelGcode(gcodePath, levelPath, true);
                }
                else
                {
                    Console.WriteLine("No levelling data found!");
                    Console.WriteLine("Please include levelling data with gcode");
                }
                originalLength = gcodeEdit.showStats(gcodePath);
                printHelp();
            }
            if (key == 'o') //Optimize holes with divide & conquer
            {
                Console.ForegroundColor = ConsoleColor.Red;
                if (!gcodeEdit.saveUndoPoint(gcodePath)) Console.WriteLine("UNABLE TO CREATE SAVE POINT!\nContinuing without saving");
                Console.ForegroundColor = ConsoleColor.White;
                gcodeEdit.optimizeHoles(gcodePath);
                originalLength = gcodeEdit.showStats(gcodePath);
                printHelp();
            }
            else if (key == 'g')    //Optimize holes with genetic algorithm
            {
                Console.ForegroundColor = ConsoleColor.Red;
                if (!gcodeEdit.saveUndoPoint(gcodePath)) Console.WriteLine("UNABLE TO CREATE SAVE POINT!\nContinuing without saving");
                Console.ForegroundColor = ConsoleColor.White;
                gcodeEdit.GA_TSP.sortGA(gcodePath, originalLength);
                originalLength = gcodeEdit.showStats(gcodePath);
                printHelp();
            }
            else if (key == 'b')
            {
                Console.ForegroundColor = ConsoleColor.Red;
                if (!gcodeEdit.saveUndoPoint(gcodePath)) Console.WriteLine("UNABLE TO CREATE SAVE POINT!\nContinuing without saving");
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write("\n\nEnter backlash in mm: ");
                double x = double.Parse(Console.ReadLine());
                gcodeEdit.correctZBacklash(gcodePath, x);
                originalLength = gcodeEdit.showStats(gcodePath);
                printHelp();
            }
            else if (key == 'd')    //Draw gcode
            {
                Form1.runForm(gcodePath);
                Console.WriteLine();
            }
            else if (key == 'u')    //Undo last action
            {
                Console.ForegroundColor = ConsoleColor.Red;
                if (gcodeEdit.undo(gcodePath)) Console.WriteLine("\n\nUndo Successful\n");
                else Console.WriteLine("\n\nUnable To Undo\n");
                Console.ForegroundColor = ConsoleColor.White;
                originalLength = gcodeEdit.showStats(gcodePath);
                printHelp();
            }
            else if (key == 'z')    //Undo last action
            {
                toolkitOptions();
                printHelp();
            }
            else if (key >= 'a' && key <= 'Z') 
            {
                Console.WriteLine(" -> Invalid key\n");
            }
        }
        Form1.shutDownForms();
        File.Delete(gcodePath.ToLower().Replace(".gcode", "_bak.gcode"));
    }

    static void printHelp()
    {
        Console.WriteLine("\n#####################################################################\n");
        Console.WriteLine("Press c to copy original gcode path");
        Console.WriteLine("Press e to optimize ETCH path");
        Console.WriteLine("Press k to level ETCH path (bilinear interpolation)");
        Console.WriteLine("Press l to level ETCH path (bicubic spline interpolation)");
        Console.WriteLine("Press o to optimize HOLE DRILL path (divide & conquer algorithm)");
        Console.WriteLine("Press g to optimize HOLE DRILL path (genetic algorithm)");
        Console.WriteLine("Press b to add backlash compensation");
        Console.WriteLine("Press d to draw gcode path");
        Console.WriteLine("Press u to undo previous action");
        Console.WriteLine("Press z to enter OPTIONS");
        Console.WriteLine("Press escape to exit");
    }
    static void toolkitOptions()
    {
        char key = 'a';
        while (key >= 'A' && key <= 'z')
        {
            printOptionsHelp();
            key = Console.ReadKey().KeyChar;
            if (key == 'a')
            {
                Console.WriteLine("\nEnter bicubic expansion grid distance(mm):");
                double x = double.Parse(Console.ReadLine());
                gcodeEdit.setBicubicExpansionGridDist(x);
            }
            else if (key >= 'a' && key <= 'Z')
            {
                Console.WriteLine(" -> Invalid key\n");
            }
        }
    }
    static void printOptionsHelp()
    {
        Console.WriteLine("\n#####################################################################\n");
        Console.WriteLine("Press a to edit bicubic spline interpolation expansion grid distance");
        Console.WriteLine("Press escape to return to main menu");
    }
}

class gcodeEdit {

    public static string[] enrichGcode(string gcodePath)
    {
        double x = 0;
        double y = 0;
        double z = 0;
        int gType = 0;
        double[] feedRate = { 60, 30 };
        string[] gcode = File.ReadAllLines(gcodePath, Encoding.UTF8);
        string[] gcodeOut = new string[gcode.Length];
        using (StreamWriter objFile = File.CreateText(gcodePath.Replace(".gcode", ".tmp")))
        {
            for (int h = 0; h < gcode.Length; h++)
            {
                string gc = gcode[h];
                if ((gc + "   ").Substring(0, 3) != "G00" && (gc + "   ").Substring(0, 3) != "G01" && (gc + "   ").Substring(0, 3) != "G0 " && (gc + "   ").Substring(0, 3) != "G1 ") gcodeOut[h] = gc;
                else
                {
                    gc = gc.Replace("G00", "G0");
                    gc = gc.Replace("G01", "G1");
                    gc = gc.Replace("\n", "");
                    string[] g = gc.Split(' ');
                    for (int j = 0; j < g.Length; j++)
                    {
                        if (g[j].Length > 0)
                        {
                            g[j] = g[j].Replace(" ", "");
                            if (g[j].Substring(0, 1) == "G") gType = int.Parse(g[j].Replace("G", ""));
                            else if (g[j].Substring(0, 1) == "X") x = double.Parse(g[j].Replace("X", ""));
                            else if (g[j].Substring(0, 1) == "Y") y = double.Parse(g[j].Replace("Y", ""));
                            else if (g[j].Substring(0, 1) == "Z") z = double.Parse(g[j].Replace("Z", ""));
                            else if (g[j].Substring(0, 1) == "F") feedRate[gType] = double.Parse(g[j].Replace("F", ""));
                        }
                    }
                    gcodeOut[h] = "G" + gType.ToString() + " X" + Math.Round(x, 2).ToString() + " Y" + Math.Round(y, 2).ToString() + " Z" + Math.Round(z, 3).ToString() + " F" + Math.Round(feedRate[gType], 2).ToString();
                }
            }
        }
        return gcodeOut;
    }

    public static double[] xyzFROMstring(string gcode)
    {
        double[] xyz = { double.NaN, double.NaN, double.NaN };
        string[] ig = gcode.Split(' ');
        for (int k = 1; k < ig.Length; k++)
        {
            if (ig[k].Length > 0)
            {
                ig[k] = ig[k].Replace(" ", "");
                char firstChar = (ig[k])[0];
                if (firstChar == 'X') xyz[0] = double.Parse(ig[k].Replace("X", ""));
                else if (firstChar == 'Y') xyz[1] = double.Parse(ig[k].Replace("Y", ""));
                else if (firstChar == 'Z') xyz[2] = double.Parse(ig[k].Replace("Z", ""));
            }
        }
        return xyz;
    }

    public static double[,] gcodeTOdouble(string[] gcode)
    {
        double[,] xyzNew = new double[gcode.Length, 3];
        double[] xyzThis = new double[3];
        for (int i = 0; i < gcode.Length; i++)
        {
            xyzThis = gcodeEdit.xyzFROMstring(gcode[i]);
            xyzNew[i, 0] = xyzThis[0];
            xyzNew[i, 1] = xyzThis[1];
            xyzNew[i, 2] = xyzThis[2];
        }
        return xyzNew;
    }



    public static double bicubicExpansionGridDist = 0.05;
    static public void setBicubicExpansionGridDist(double dist)
    {
        if (dist > 0) bicubicExpansionGridDist = dist;
    }
    static public bool saveUndoPoint(string gcodePath)
    {
        string backupPath = gcodePath.ToLower().Replace(".gcode", "_bak.gcode");
        bool fileExists = File.Exists(backupPath);
        try
        {
            string backupGcode = "";
            if (fileExists)
            {
                backupGcode = File.ReadAllText(backupPath);
                File.Delete(backupPath);
            }
            using (StreamWriter objFile = File.CreateText(backupPath))
            {
                objFile.WriteLine("----beginBackup----");
                string gcode = File.ReadAllText(gcodePath);
                objFile.Write(gcode);
                objFile.WriteLine("----endBackup----");
                if (fileExists) objFile.Write(backupGcode);
            }
        }
        catch
        {
            return false;
        }
        return true;
    }

    static public bool undo(string gcodePath)
    {
        try
        {
            string backupPath = gcodePath.ToLower().Replace(".gcode", "_bak.gcode");
            if (!File.Exists(backupPath)) return false;
            string[] backupGcode = File.ReadAllLines(backupPath, Encoding.UTF8);
            int i = 0;
            while (i < backupGcode.Length && backupGcode[i] != "----beginBackup----") i++;
            if (i == backupGcode.Length) return false;
            i++;
            File.Delete(gcodePath);
            using (StreamWriter objFile = File.CreateText(gcodePath))
            {
                while (i < backupGcode.Length && backupGcode[i] != "----endBackup----")
                {
                    objFile.WriteLine(backupGcode[i]);
                    i++;
                }
                File.Delete(backupPath);
            }
            i++;
            if (i >= backupGcode.Length) File.Delete(backupPath);
            else
            {
                using (StreamWriter objFile = File.CreateText(backupPath))
                {
                    while (i < backupGcode.Length)
                    {
                        objFile.WriteLine(backupGcode[i]);
                        i++;
                    }
                }
            }
        }
        catch
        {
            return false;
        }
        return true;
    }


    static public void copyGcode(string gcodePath, double x, double y)
    {
        string[] igcode = File.ReadAllLines(gcodePath, Encoding.UTF8);
        int i = 0;
        int oLength = 0;
        using (StreamWriter objFile = File.CreateText(gcodePath.Replace(".gcode", ".tmp")))
        {
            double lastZ = 0;
            while (i < igcode.Length)
            {
                if (igcode[i] == ";End of original gcode") oLength = i;
                objFile.WriteLine(igcode[i]);
                string[] g = igcode[i].Split(' ');
                for (int j = 0; j < g.Length; j++)
                {
                    if (g[j].Length > 1)
                    {
                        string axis = (g[j]).Substring(0, 1);
                        if (axis == "z" || axis == "Z")
                        {
                            try
                            {
                                lastZ = double.Parse(g[j].Substring(1));
                            }
                            catch { }
                        }
                    }
                }
                i++;
            }
            if (oLength == 0)
            {
                objFile.WriteLine(";End of original gcode");
                oLength = igcode.Length;
            }
            if (lastZ < 0.1) objFile.WriteLine("G0 Z1.5");
            i = 0;
            while (i < oLength)
            {
                string[] g = igcode[i].Split(' ');
                bool lineAltered = false;
                for (int j = 0; j < g.Length; j++)
                {
                    if (g[j].Length > 1)
                    {
                        string axis = (g[j] + " ").Substring(0, 1);
                        if (axis == ";") j = g.Length;
                        else if (axis == "x" || axis == "X")
                        {
                            try
                            {
                                double thisx = double.Parse(g[j].Substring(1));
                                thisx += x;
                                g[j] = "X" + Math.Round(thisx, 3).ToString();
                                lineAltered = true;
                            }
                            catch
                            {
                                Console.WriteLine("ERROR");
                            }
                        }
                        else if (axis == "y" || axis == "Y")
                        {
                            try
                            {
                                double thisy = double.Parse(g[j].Substring(1));
                                thisy += y;
                                g[j] = "Y" + Math.Round(thisy, 3).ToString();
                                lineAltered = true;
                            }
                            catch
                            {
                                Console.WriteLine("ERROR");
                            }
                        }
                    }
                }
                if (!lineAltered) objFile.WriteLine(igcode[i]);
                else
                {
                    for (int j = 0; j < g.Length; j++)
                    {
                        objFile.Write(g[j]);
                        if (j < g.Length - 1) objFile.Write(" ");
                    }
                    objFile.WriteLine();
                }
                i++;
            }
        }
        //Copy data from .tmp to .gcode and add header info
        string currentContent = File.ReadAllText(gcodePath.Replace(".gcode", ".tmp"));
        if (System.IO.File.Exists(gcodePath.Replace(".tmp", ".gcode"))) File.Delete(gcodePath.Replace(".tmp", ".gcode")); //try/catch exception handling needs to be implemented

        File.WriteAllText(gcodePath, ";Path copied (X" + x.ToString() + " Y" + y.ToString() + ") with td0g's PCB Gcode Toolkit v" + Program.currentVersion + "\n;See www.github.com/td0g \n\n" + currentContent);
        System.IO.File.Delete(gcodePath.Replace(".gcode", ".tmp"));
    }



    static public void levelGcode(string oName, string levelPath, bool bicubic)
    {
        string[] gcode = File.ReadAllLines(oName, Encoding.UTF8);
        int decimalPlaces = 2;
        double ignoreAboveZ = 1;
        double xMax = -9999999;
        double xMin = 9999999;
        double yMax = -9999999;
        double yMin = 9999999;
        double zMax = -9999999;
        double zMin = 9999999;
        double zMaxLevelled = -9999999;
        double zMinLevelled = 9999999;
        string[] level = File.ReadAllLines(levelPath, Encoding.UTF8);
        double[] yList = new double[100]; //The yList will be sorted in ascending order.  y=0 is not necesarily yList(0) or yList(yListMax)
        double[] xList = new double[100];
        int yListSize = 0;
        int xListSize = 0;
        double zZero = 0;
        double distZero = 9999999;
        string gType = "";
        string feedRateA = "F120";
        string feedRateB = "F45";

        //Make xyz list
        for (int i = 0; i < level.Length; i++)
        {
            string c = level[i];
            char[] delimiter = { ' ', ',', '/' };
            string[] b = c.Split(delimiter);
            if (b.Length == 3)
            {
                try
                {
                    float x = float.Parse(b[0]);
                    float y = float.Parse(b[1]);
                    float z = float.Parse(b[2]);
                    if (Math.Sqrt(x * x + y * y) < distZero)
                    {
                        zZero = z;
                        distZero = Math.Sqrt(x * x + y * y);
                    }
                    bool foun = false;
                    for (int j = 0; j < yListSize; j++)
                    {
                        if (yList[j] == y)
                        {
                            foun = true;
                            j = yListSize;

                        }
                    }
                    if (!foun)
                    {
                        yList[yListSize] = y;
                        yListSize++;
                    }
                    foun = false;
                    for (int j = 0; j < xListSize; j++)
                    {
                        if (xList[j] == x)
                        {
                            foun = true;
                            j = xListSize;
                        }
                    }
                    if (!foun)
                    {
                        xList[xListSize] = x;
                        xListSize++;

                    }
                }
                catch { }
            }
        }
        Array.Resize<double>(ref xList, xListSize);
        Array.Resize<double>(ref yList, yListSize);
        //Sort x and y
        xListSize--;
        yListSize--;

        for (int i = xListSize; i > -1; i--)
        {
            for (int j = 0; j < i; j++)
            {
                if (xList[j] > xList[j + 1])
                {
                    double temp = xList[j + 1];
                    xList[j + 1] = xList[j];
                    xList[j] = temp;
                }
            }
        }
        for (int i = yListSize; i > -1; i--)
        {
            for (int j = 0; j < i; j++)
            {
                if (yList[j] > yList[j + 1])
                {
                    double temp = yList[j + 1];
                    yList[j + 1] = yList[j];
                    yList[j] = temp;
                }
            }
        }
        //########################################################################################
        //                          Z Array Table
        //########################################################################################

        //Build empty Z array
        double[,] zArray = new double[xListSize + 1, yListSize + 1];
        for (int i = 0; i <= xListSize; i++)
        {
            for (int j = 0; j <= yListSize; j++) zArray[i, j] = 1000;
        }

        //Populate zArray
        for (int i = 0; i < level.Length; i++)
        {
            string c = level[i];
            char[] delimiter = { ' ', ',', '/' };
            string[] b = c.Split(delimiter);
            float x = 9999999;
            float y = 9999999;
            float z = 0;
            if (b.Length == 3)
            {
                try
                {
                    x = float.Parse(b[0]);
                }
                catch { }
                try
                {
                    y = float.Parse(b[1]);
                }
                catch { }
                try
                {
                    z = float.Parse(b[2]);
                }
                catch { }
                for (int j = 0; j <= xListSize; j++)
                {
                    for (int k = 0; k <= yListSize; k++)
                    {
                        if (x == xList[j] && y == yList[k])
                        {
                            zArray[j, k] = Math.Round(z - zZero, decimalPlaces);
                            if (zArray[j, k] > zMaxLevelled) zMaxLevelled = zArray[j, k];
                            else if (zArray[j, k] < zMinLevelled) zMinLevelled = zArray[j, k];
                            k = yListSize + 1;
                            j = xListSize + 1;
                        }
                    }
                }
            }
        }

        //Check for blank entries in zArray, exit script if any found
        int emptyZValues = 0;
        for (int i = 0; i <= xListSize; i++)
        {
            for (int j = 0; j <= yListSize; j++)
            {
                if (zArray[i, j] > 999)
                {
                    emptyZValues++;
                    double[,] yzArray = new double[yListSize,2];
                    int k = 0;
                    for (int l = 0; l <= yListSize; l++)
                    {
                        if (l != j)
                        {
                            yzArray[k, 0] = yList[l];
                            yzArray[k, 1] = zArray[i, l];
                            k++;
                        }

                    }
                    double[,] xzArray = new double[xListSize, 2];
                    k = 0;
                    for (int l = 0; l <= xListSize; l++)
                    {
                        if (l != i)
                        {
                            xzArray[k, 0] = xList[l];
                            xzArray[k, 1] = zArray[l, j];
                            k++;
                        }

                    }
                    double[,] xzExpandedThis = SplineInterpolate.bci(xzArray);
                    double[,] yzExpandedThis = SplineInterpolate.bci(yzArray);
                    double xzDist = 9999;
                    double yzDist = 9999;
                    double newZx = 0;
                    double newZy = 0;
                    for (k = 0; k < xzExpandedThis.Length/2; k++)
                    {
                        if (Math.Abs(xzExpandedThis[k, 0] - xList[i]) < xzDist)
                        {
                            xzDist = Math.Abs(xzExpandedThis[k, 0] - xList[i]);
                            newZx = xzExpandedThis[k, 1];
                        }
                    }
                    for (k = 0; k < yzExpandedThis.Length/2; k++)
                    {
                        if (Math.Abs(yzExpandedThis[k, 0] - yList[j]) < yzDist)
                        {
                            yzDist = Math.Abs(yzExpandedThis[k, 0] - yList[j]);
                            newZy = yzExpandedThis[k, 1];
                        }
                    }
                    zArray[i, j] = (newZy + newZx) / 2;
                }
            }
        }
        Console.WriteLine("\n\nLevelled data parsed");
        if (emptyZValues > 0)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("\n### WARNINING ###");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("Missing Probe Locations Found: " + emptyZValues.ToString() + "\nScript will attempt to produce Gcode anyway\nIt is recommended not to use output");
            Console.WriteLine();
        }

        //########################################################################################
        //                          Write X & Y Lists, Z Array to CSV File
        //########################################################################################

        using (StreamWriter objFile = File.CreateText((levelPath.ToLower()).Replace(".txt", "_array.csv")))
        {
            Console.Write("   Z-PROBE HEAT MAP   ");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("Highest Z Probe: ");
            Console.Write(zMaxLevelled);
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.Write("   Lowest Z Probe: ");
            Console.WriteLine(zMinLevelled);
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.White;
            objFile.Write(",");
            for (int i = 0; i <= xListSize; i++)
            {
                objFile.Write(Math.Round(xList[i], 3).ToString() + ",");
            }
            objFile.WriteLine();
            for (int i = 0; i <= yListSize; i++)
            {
                objFile.Write(Math.Round(yList[yListSize - i], 3).ToString());
                for (int j = 0; j <= xListSize; j++)
                {
                    objFile.Write(",");
                    objFile.Write(Math.Round(zArray[j, yListSize - i], 3).ToString());
                    double thisZ = zArray[j, yListSize - i];
                    if (thisZ < (zMinLevelled + (zMaxLevelled - zMinLevelled) * 0.14)) Console.ForegroundColor = ConsoleColor.Black;
                    else if (thisZ < (zMinLevelled + (zMaxLevelled - zMinLevelled) * 0.29)) Console.ForegroundColor = ConsoleColor.DarkBlue;
                    else if (thisZ < (zMinLevelled + (zMaxLevelled - zMinLevelled) * 0.42)) Console.ForegroundColor = ConsoleColor.Blue;
                    else if (thisZ < (zMinLevelled + (zMaxLevelled - zMinLevelled) * 0.57)) Console.ForegroundColor = ConsoleColor.White;
                    else if (thisZ < (zMinLevelled + (zMaxLevelled - zMinLevelled) * 0.71)) Console.ForegroundColor = ConsoleColor.Red;
                    else if (thisZ < (zMinLevelled + (zMaxLevelled - zMinLevelled) * 0.85)) Console.ForegroundColor = ConsoleColor.DarkRed;
                    else Console.ForegroundColor = ConsoleColor.Black;
                    Console.Write("#");
                }
                Console.WriteLine();
                objFile.WriteLine();
            }
        }
        Console.ForegroundColor = ConsoleColor.White;
        Console.WriteLine("\n\nArray file: " + (levelPath.ToLower()).Replace(".txt", "_array.csv"));

        //########################################################################################

        //                          Interpolate XYZ data

        //########################################################################################

        if (bicubic)
        {
            Console.WriteLine("\nPerforming bicubic spline interpolation to improve accuracy\n");
            SplineInterpolate.expandZList(zArray, xList, yList, (levelPath.ToLower()).Replace(".txt", "_arrayExp.csv"));
        }
        //########################################################################################
        //                          Parse all Gcode
        //########################################################################################

        double[,] hole = new double[gcode.Length, 2];
        double xLast = 0;
        double yLast = 0;
        double xIntPos = 0.0;
        double yIntPos = 0.0;
        double zIntPos = 0.0;
        double xTarget = 0.0;
        double yTarget = 0.0;
        double zTarget = 0.0;
        double xLastIntPos = 0.0;
        double yLastIntPos = 0.0;
        double zLastIntPos = 0.0;
        double xLastTarget = 0.0;
        double yLastTarget = 0.0;
        double zLastTarget = 0.0;
        double zRoundedFloatOld = 999.999;

        bool firstMove = true;
        
        string lastFRA = "";
        string lastFRB = "";

        int newTotalCount = 0;
        int oldTotalCount = 0;
        int cornerCount = 0;
        int edgeCount = 0;


        //Get starting Z position
        for (int i = 0; i <= gcode.Length; i++)
        {
            string gc = gcode[i];

            if ((gc + "   ").Substring(0, 3) == "G00" || (gc + "   ").Substring(0, 3) == "G01" || (gc + "   ").Substring(0, 3) == "G0 " || (gc + "   ").Substring(0, 3) == "G1 ")
            {
                string[] g = gc.Split(' ');
                for (int j = 0; j < g.Length - 1; j++)
                {
                    g[j] = g[j].Replace(" ", "");
                    if (g[j].Substring(0, 1) == "Z")
                    {
                        zLastTarget = float.Parse(g[j].Replace("Z", ""));
                        zTarget = zLastTarget;
                        j = g.Length + 1;
                        i = gcode.Length + 1;
                    }
                }
            }
        }

        using (StreamWriter objFile = File.CreateText(oName.Replace(".gcode", ".tmp")))
        {
            for (int h = 0; h < gcode.Length; h++)
            {
                xLast = xTarget;
                yLast = yTarget;
                string gc = gcode[h];
                bool x = false;
                bool y = false;
                bool z = false;
                if ((gc + "   ").Substring(0, 3) != "G00" && (gc + "   ").Substring(0, 3) != "G01" && (gc + "   ").Substring(0, 3) != "G0 " && (gc + "   ").Substring(0, 3) != "G1 ") objFile.WriteLine(gc);
                else if (gc.Substring(0, 1) == "G")
                {
                    oldTotalCount++;
                    gc = gc.Replace("G00", "G0");
                    gc = gc.Replace("G01", "G1");
                    gc = gc.Replace("\n", "");
                    gc = gc.Replace("\n", "");
                    string[] g = gc.Split(' ');
                    xLastTarget = xTarget;
                    yLastTarget = yTarget;
                    zLastTarget = (float)zTarget;
                    gType = g[0];
                    for (int j = 1; j < g.Length; j++)
                    {
                        if (g[j].Length > 0)
                        {
                            g[j] = g[j].Replace(" ", "");
                            if (g[j].Substring(0, 1) == "X")
                            {
                                xTarget = float.Parse(g[j].Replace("X", ""));
                                if (xTarget != xLastTarget || firstMove) x = true;
                                if (xTarget > xMax) xMax = xTarget;
                                if (xTarget < xMin) xMin = xTarget;
                            }
                            else if (g[j].Substring(0, 1) == "Y")
                            {
                                yTarget = float.Parse(g[j].Replace("Y", ""));
                                if (yTarget != yLastTarget || firstMove) y = true;
                                if (yTarget > yMax) yMax = yTarget;
                                if (yTarget < yMin) yMin = yTarget;
                            }
                            else if (g[j].Substring(0, 1) == "Z")
                            {
                                zTarget = float.Parse(g[j].Replace("Z", ""));
                                if (zTarget != zLastTarget || firstMove) z = true;
                                if (zTarget > zMax) zMax = zTarget;
                                if (zTarget < zMin) zMin = zTarget;
                            }
                            else if (g[j].Substring(0, 1) == "F")
                            {
                                if (gType == "G0") feedRateA = g[j];
                                else feedRateB = g[j];
                            }
                        }
                    }


                    //########################################################################################
                    //                          'Calculate position of intermediate point
                    //########################################################################################

                    if (x || y || z || firstMove)
                    {
                        double dist = Math.Sqrt((xTarget - xLastTarget) * (xTarget - xLastTarget) + (yTarget - yLastTarget) * (yTarget - yLastTarget) + (zTarget - zLastTarget) * (zTarget - zLastTarget));
                        int seg = -1;
                        int numSegs = 0;
                        double[,] bicubicPath = { { 0 } };
                        double P = 0;
                        while ((xIntPos != xTarget || yIntPos != yTarget || zIntPos != zTarget || firstMove) && seg != numSegs)
                        {
                            int xLine = 0;
                            int yLine = 0;
                            xLastIntPos = xIntPos;
                            yLastIntPos = yIntPos;
                            zLastIntPos = zIntPos;
                            if (!x && !y && !firstMove) //Only moving Z - go straight there
                            {
                                if (bicubic && zTarget < 0.1)
                                {
                                    zTarget += SplineInterpolate.getNearest(xIntPos, yIntPos);
                                }
                                zIntPos = zTarget;
                            }
                            else if (zLastTarget > ignoreAboveZ || zTarget > ignoreAboveZ || firstMove) //We are above workpiece or this is the first move, just make one line
                            {
                                firstMove = false;
                                xIntPos = xTarget;
                                yIntPos = yTarget;
                                zIntPos = zTarget;
                            }
                            //Divide into segments
                            else if (bicubic)
                            {
                                seg++;
                                if (seg == 0) {
                                    double[,] twoPoints = new double[2, 3];
                                    twoPoints[0, 0] = xLastTarget;
                                    twoPoints[0, 1] = yLastTarget;
                                    twoPoints[0, 2] = zLastTarget;
                                    twoPoints[1, 0] = xTarget;
                                    twoPoints[1, 1] = yTarget;
                                    twoPoints[1, 2] = zTarget;
                                    bicubicPath = gcodeEdit.splineToLines(twoPoints);
                                    numSegs = bicubicPath.GetLength(0)-1;
                                    seg++;
                                }
                                xIntPos = bicubicPath[seg, 0];
                                yIntPos = bicubicPath[seg, 1];
                                zIntPos = zLastTarget + (zTarget - zLastTarget) * seg / numSegs + bicubicPath[seg, 2];
                            }				  
                            //1. Are we going to cross an X or Y int before reaching target?
                            //2. Which line are we going to cross first?
                            else
                            {
                                if (x)
                                {
                                    if (xLastIntPos < xList[0] || xTarget <= xList[0])
                                    {
                                        //'objFile.Write("A ");
                                        if (xLastIntPos <= xList[0] && xTarget <= xList[0]) xIntPos = xTarget;
                                        else xIntPos = xList[0];
                                    }
                                    else if (xLastIntPos > xList[xListSize] || xTarget >= xList[xListSize])
                                    {
                                        if (xLastIntPos >= xList[xListSize] && xTarget >= xList[xListSize]) xIntPos = xTarget;
                                        else xIntPos = xList[xListSize];
                                        xLine = xListSize;
                                    }
                                    else
                                    {
                                        if (xTarget > xLastTarget)
                                        {
                                            xLine = 0;
                                            while (xIntPos >= xList[xLine]) xLine++;
                                        }
                                        else
                                        {
                                            xLine = xListSize;
                                            while (xIntPos <= xList[xLine]) xLine--;
                                        }
                                        if (xTarget > xIntPos)
                                        {
                                            if (xTarget > xList[xLine]) xIntPos = xList[xLine];
                                            else xIntPos = xTarget;
                                        }
                                        else if (xTarget < xIntPos)
                                        {
                                            if (xTarget < xList[xLine]) xIntPos = xList[xLine];
                                            else xIntPos = xTarget;
                                        }
                                    }
                                }
                                else xIntPos = xTarget; //X is not moving, just set to final target


                                
                                if (y)  //Quickly find intermediate position for Y (will sort out X/Y disputes later)
                                {
                                    if (yLastIntPos < yList[0] || yTarget <= yList[0])  //We are starting below probe grid OR finishing below probe grid
                                    {
                                        if (yLastIntPos <= yList[0] && yTarget <= yList[0]) yIntPos = yTarget;  //Starting AND finishing below grid - move in straight line without intermediate positions
                                        else yIntPos = yList[0];    //Starting OR finishing within grid - move to edge of grid
                                    }
                                    else if (yLastIntPos > yList[yListSize] || yTarget >= yList[yListSize])  //We are starting above probe grid OR finishing above probe grid
                                    {
                                        if (yLastIntPos >= yList[yListSize] && yTarget >= yList[yListSize]) yIntPos = yTarget;  //Starting AND finishing above grid - move in straight line without intermediate positions
                                        else yIntPos = yList[yListSize];    //Starting OR finishing within grid - move to edge of grid
                                    }
                                    else   //We are starting AND finishing within probe grid
                                    {
                                        if (yTarget > yLastTarget)
                                        {
                                            yLine = 0;
                                            while (yIntPos >= yList[yLine]) yLine++;
                                        }
                                        else
                                        {
                                            yLine = yListSize;
                                            while (yIntPos <= yList[yLine]) yLine--;
                                        }
                                        if (yTarget > yIntPos)  //Moving upward
                                        {
                                            if (yTarget > yList[yLine]) yIntPos = yList[yLine];
                                            else yIntPos = yTarget;
                                        }
                                        else if (yTarget < yIntPos) //Moving downward
                                        {
                                            if (yTarget < yList[yLine]) yIntPos = yList[yLine];
                                            else yIntPos = yTarget;
                                        }
                                    }
                                }
                                else yIntPos = yTarget; //y is not moving, just set to final target

                                //What intPos comes first?
                                if (!y)
                                {   //Y is not moving, X comes first
                                    yIntPos = yLastTarget;
                                    zIntPos = (xIntPos - xLastTarget) / (xTarget - xLastTarget) * (zTarget - zLastTarget) + zLastTarget;
                                }

                                else if (!x)
                                {   //X is not moving, Y comes first
                                    xIntPos = xLastTarget;
                                    zIntPos = (yIntPos - yLastTarget) / (yTarget - yLastTarget) * (zTarget - zLastTarget) + zLastTarget;
                                }

                                else if ((xIntPos - xLastIntPos) / (xTarget - xLastTarget) > (yIntPos - yLastIntPos) / (yTarget - yLastTarget))
                                { //Crossing Y line first
                                    xIntPos = (yIntPos - yLastTarget) / (yTarget - yLastTarget) * (xTarget - xLastTarget) + xLastTarget;
                                    zIntPos = (yIntPos - yLastTarget) / (yTarget - yLastTarget) * (zTarget - zLastTarget) + zLastTarget;
                                }
                                else
                                { //Crossing X line first
                                    yIntPos = (xIntPos - xLastTarget) / (xTarget - xLastTarget) * (yTarget - yLastTarget) + yLastTarget;
                                    zIntPos = (xIntPos - xLastTarget) / (xTarget - xLastTarget) * (zTarget - zLastTarget) + zLastTarget;
                                }
                            }

                            //########################################################################################
                            //                          'Apply Bilinear Interpolation
                            //########################################################################################
                            //Is this a hop?
                            if (!bicubic)
                            {
                                if (zIntPos > ignoreAboveZ)
                                {
                                    P = 0;
                                }
                                //Is point outside corner?
                                else if (xIntPos <= xList[0] && yIntPos <= yList[0])
                                {
                                    P = zArray[0, 0];
                                    cornerCount++;
                                }
                                else if (xIntPos <= xList[0] && yIntPos >= yList[yListSize])
                                {
                                    P = zArray[0, yListSize];
                                    cornerCount++;
                                }
                                else if (xIntPos >= xList[xListSize] && yIntPos >= yList[yListSize])
                                {
                                    P = zArray[xListSize, yListSize];
                                    cornerCount++;
                                }
                                else if (xIntPos >= xList[xListSize] && yIntPos <= yList[0])
                                {
                                    P = zArray[xListSize, 0];
                                    cornerCount++;
                                }
                                //Is point outside edge?
                                else if (xIntPos <= xList[0])
                                {
                                    int yP = 0;
                                    edgeCount++;
                                    for (int j = 0; j <= yListSize; j++)
                                    {
                                        if (yIntPos < yList[j])
                                        {
                                            yP = j;
                                            j = yListSize + 1;
                                        }
                                    }
                                    double yA = yList[yP - 1];
                                    double yB = yList[yP];
                                    double zAB = zArray[0, yP - 1];
                                    double zBB = zArray[0, yP];
                                    P = ((yB - yIntPos) / (yB - yA)) * zAB + ((yIntPos - yA) / (yB - yA)) * zBB;
                                }
                                else if (xIntPos >= xList[xListSize])
                                {
                                    int yP = 0;
                                    edgeCount++;
                                    for (int j = 0; j <= yListSize; j++)
                                    {
                                        if (yIntPos < yList[j])
                                        {
                                            yP = j;
                                            j = yListSize + 1;
                                        }
                                    }
                                    double yA = yList[yP - 1];
                                    double yB = yList[yP];
                                    double zAB = zArray[xListSize, yP - 1];
                                    double zBB = zArray[xListSize, yP];
                                    P = ((yB - yIntPos) / (yB - yA)) * zAB + ((yIntPos - yA) / (yB - yA)) * zBB;
                                }

                                else if (yIntPos <= yList[0])
                                {
                                    edgeCount++;
                                    int xP = 0;
                                    for (int j = 0; j <= xListSize; j++)
                                    {
                                        if (xIntPos < xList[j])
                                        {
                                            xP = j;
                                            j = xListSize + 1;
                                        }
                                    }
                                    double xA = xList[xP - 1];
                                    double xB = xList[xP];
                                    double zAA = zArray[xP - 1, 0];
                                    double zBA = zArray[xP, 0];
                                    P = ((xIntPos - xA) / (xB - xA)) * zBA + ((xB - xIntPos) / (xB - xA)) * zAA;
                                }
                                else if (yIntPos >= yList[yListSize])
                                {
                                    edgeCount++;
                                    int xP = 0;
                                    for (int j = 0; j <= xListSize; j++)
                                    {
                                        if (xIntPos < xList[j])
                                        {
                                            xP = j;
                                            j = xListSize + 1;
                                        }
                                    }
                                    double xA = xList[xP - 1];
                                    double xB = xList[xP];
                                    double zAA = zArray[xP - 1, yListSize];
                                    double zBA = zArray[xP, yListSize];
                                    P = ((xIntPos - xA) / (xB - xA)) * zBA + ((xB - xIntPos) / (xB - xA)) * zAA;
                                }

                                //Inside Grid
                                else
                                {
                                    if (yIntPos == yList[yLine])
                                    {
                                        if (xIntPos == xList[xLine])
                                        {
                                            P = zArray[xLine, yLine];
                                        }
                                        else
                                        {
                                            int xP = 0;
                                            for (int j = 0; j <= xListSize; j++)
                                            {
                                                if (xIntPos < xList[j])
                                                {
                                                    xP = j;
                                                    j = xListSize + 1;
                                                }
                                            }
                                            double xA = xList[xP - 1];
                                            double xB = xList[xP];
                                            double zAA = zArray[xP - 1, yLine];
                                            double zBA = zArray[xP, yLine];
                                            P = ((xIntPos - xA) / (xB - xA)) * zBA + ((xB - xIntPos) / (xB - xA)) * zAA;
                                        }

                                    }
                                    else if (xIntPos == xList[xLine])
                                    {
                                        int yP = 0;
                                        edgeCount++;
                                        for (int j = 0; j <= yListSize; j++)
                                        {
                                            if (yIntPos < yList[j])
                                            {
                                                yP = j;
                                                j = yListSize + 1;
                                            }
                                        }
                                        double yA = yList[yP - 1];
                                        double yB = yList[yP];
                                        double zAB = zArray[xLine, yP - 1];
                                        double zBB = zArray[xLine, yP];
                                        P = ((yB - yIntPos) / (yB - yA)) * zAB + ((yIntPos - yA) / (yB - yA)) * zBB;
                                    }
                                    else
                                    {
                                        //--> Bilinear Interpolation
                                        // https://en.wikipedia.org/wiki/Bilinear_interpolation
                                        //http://supercomputingblog.com/graphics/coding-bilinear-interpolation/
                                        int xP = 0;
                                        int yP = 0;
                                        for (int j = 0; j <= xListSize; j++)
                                        {
                                            if (xIntPos < xList[j])
                                            {
                                                xP = j;
                                                j = xListSize + 1;
                                            }
                                        }
                                        for (int j = 0; j <= yListSize; j++)
                                        {
                                            if (yIntPos < yList[j])
                                            {
                                                yP = j;
                                                j = yListSize + 1;
                                            }
                                        }
                                        double xA = xList[xP - 1];
                                        double xB = xList[xP];
                                        double yA = yList[yP - 1];
                                        double yB = yList[yP];
                                        double zAA = zArray[xP - 1, yP - 1];
                                        double zBA = zArray[xP - 1, yP];
                                        double zAB = zArray[xP, yP - 1];
                                        double zBB = zArray[xP, yP];
                                        double rA = ((xIntPos - xA) / (xB - xA)) * zBA + ((xB - xIntPos) / (xB - xA)) * zAA;
                                        double rB = ((xIntPos - xA) / (xB - xA)) * zBB + ((xB - xIntPos) / (xB - xA)) * zAB;
                                        P = ((yB - yIntPos) / (yB - yA)) * rA + ((yIntPos - yA) / (yB - yA)) * rB;
                                    }
                                }
                            }


                            //########################################################################################
                            //                          'Write to Gcode File
                            //########################################################################################

                            double zRoundedFloat = Math.Round(zIntPos + P, decimalPlaces);
                            if (x || y || zRoundedFloat != zRoundedFloatOld)
                            {
                                objFile.Write(gType);
                                newTotalCount++;
                            }
                            if (seg == -1)
                            {
                                if (x)
                                {
                                    double xRoundedFloat = Math.Round(xIntPos, decimalPlaces);
                                    objFile.Write(" X");
                                    objFile.Write(xRoundedFloat);
                                }
                                if (y)
                                {
                                    double yRoundedFloat = Math.Round(yIntPos, decimalPlaces);
                                    objFile.Write(" Y");
                                    objFile.Write(yRoundedFloat);
                                }
                                if (z || x || y)
                                {
                                    zRoundedFloat = Math.Round(zIntPos + P, decimalPlaces + 1);
                                    if (zRoundedFloat != zRoundedFloatOld)
                                    {
                                        objFile.Write(" Z");
                                        objFile.Write(zRoundedFloat);
                                    }

                                }
                            }
                            else
                            {
                                objFile.Write(" X" + Math.Round(xIntPos, decimalPlaces).ToString() + " Y" + Math.Round(yIntPos, decimalPlaces).ToString());
                                if (Math.Round(zIntPos, decimalPlaces) != zRoundedFloatOld)
                                {
                                    objFile.Write(" Z" + Math.Round(zIntPos, decimalPlaces).ToString());
                                    zRoundedFloatOld = Math.Round(zIntPos, decimalPlaces);
                                }
                            }
                            if (gType == "G0" && feedRateA != lastFRA)
                            {
                                objFile.Write(" ");
                                objFile.Write(feedRateA);
                                lastFRA = feedRateA;
                            }
                            else if (gType == "G1" && feedRateB != lastFRB)
                            {
                                objFile.Write(" ");
                                objFile.Write(feedRateB);
                                lastFRB = feedRateB;
                            }
                            objFile.WriteLine();
                            zRoundedFloatOld = zRoundedFloat;
                        }
                    }
                }
            }
        }
        //########################################################################################
        //                          Wrap It Up
        //########################################################################################

        //Copy data from .tmp to .gcode and add header info
        string currentContent = File.ReadAllText(oName.Replace(".gcode", ".tmp"));
        if (System.IO.File.Exists(oName.Replace(".tmp", ".gcode"))) File.Delete(oName.Replace(".tmp", ".gcode")); //try/catch exception handling needs to be implemented

        File.WriteAllText(oName, ";Levelled with td0g's PCB Gcode Toolkit v" + Program.currentVersion + "\n;See www.github.com/td0g \n\n" + currentContent);
        System.IO.File.Delete(oName.Replace(".gcode", ".tmp"));


        //Notify User 
        float percentInside = 100 * (1 - (edgeCount + cornerCount) / oldTotalCount);
        Console.WriteLine("\nLevelled Gcode: " + oName.Replace(".tmp", ".gcode"));
        Console.WriteLine("\n\nPoints before / after:  " + oldTotalCount.ToString() + " / " + newTotalCount.ToString());
        Console.WriteLine("Points at or outside corners:  " + cornerCount.ToString() + "\nPoints at or outside edges:    " + edgeCount.ToString());
        Console.WriteLine("Percentage inside probed area:  " + Math.Round(percentInside, 1).ToString() + "%");


    }

    static public double[,] splineToLines(double[,] points)   //x,y,z
    {
        double eCrit = 0.005;
        double e = gcodeEdit.averageError(points);
        int numPoints = points.GetLength(0) - 1;
        for (int i = 0; i < numPoints; i++)
        {
            points[i, 2] = points[i, 2] + e;
        }
        e = gcodeEdit.maxError(points);
        if (e > eCrit)
        {
            int j = 0;
            int newNumPoints = (numPoints) * 2 + 1;
            double[,] newPoints = new double[newNumPoints, 3];
            while (j < newNumPoints)
            {
                if (j % 2 == 0)
                {
                    newPoints[j, 0] = points[j / 2, 0];
                    newPoints[j, 1] = points[j / 2, 1];
                    newPoints[j, 2] = points[j / 2, 2];
                }
                else
                {
                    newPoints[j, 0] = (points[j / 2, 0] + points[j / 2 + 1, 0]) / 2;
                    newPoints[j, 1] = (points[j / 2, 1] + points[j / 2 + 1, 1]) / 2;
                    newPoints[j, 2] = (points[j / 2, 1] + points[j / 2 + 1, 2]) / 2;
                }
                j++;
            }
            return splineToLines(newPoints);
        }
        return points;
    }

    static public double averageError(double[,] points)   //x, y, z
    {
        int numPoints = points.GetLength(0) - 1;

        for (int i = 0; i <= numPoints; i++)
        {
            points[i, 2] = SplineInterpolate.getNearest(points[i, 0], points[i, 1]);    //Sets Z to zero???
        }
        double totalDistSquare = Math.Pow((points[0, 0] - points[numPoints, 0]), 2) + Math.Pow((points[0, 1] - points[numPoints, 1]), 2);
        int numDivisions = Math.Min(100, (int)(totalDistSquare * 100));
        numDivisions = Math.Max(numDivisions, 5);

        numDivisions /= (numPoints);
    
    double e = 0;
        for (int i = 0; i < numPoints; i++)
        {
            for (int j = 0; j <= numDivisions; j++)
            {
                double x = points[i, 0] + (points[i + 1, 0] - points[i, 0]) / numDivisions * j;
                double y = points[i, 1] + (points[i + 1, 1] - points[i, 1]) / numDivisions * j;
                double zSpline = SplineInterpolate.getNearest(x, y);
                double zLine = points[i, 2] + (points[i + 1, 2] - points[i, 2]) / numDivisions * j;
                e = e + zLine - zSpline;
            }
        }
        e = e / (double)(numDivisions * numPoints);
        return e;
    }

    static public double maxError(double[,] points)   //x, y, z
    {
        int numPoints = points.GetLength(0) - 1;
        double totalDistSquare = Math.Pow((points[0, 0] - points[numPoints, 0]), 2) + Math.Pow((points[0, 1] - points[numPoints, 1]), 2);
        int numDivisions = Math.Min(100, (int)(totalDistSquare * 100));
        numDivisions = Math.Max(numDivisions, 5);
        numDivisions /= (numPoints);
    
    double maxE = 0;
        for (int i = 0; i < numPoints; i++)
        {
            for (int j = 0; j <= numDivisions; j++)
            {
                double x = points[i, 0] + (points[i + 1, 0] - points[i, 0]) / numDivisions * j;
                double y = points[i, 1] + (points[i + 1, 1] - points[i, 1]) / numDivisions * j;
                double zSpline = SplineInterpolate.getNearest(x, y);
                double zLine = points[i, 2] + (points[i + 1, 2] - points[i, 2]) / numDivisions * j;
                maxE = Math.Max(maxE, Math.Abs(zLine - zSpline));
            }
        }
        return maxE;
    }

    
    static public bool combineEtches(string gcodePath, double dist, bool finalSort)
    {
        Console.WriteLine("Loading gcode");
        bool success = false;
        double distSquare = dist * dist;
        string[] gcode = enrichGcode(gcodePath);
        int[,] etchStartEnd = new int[gcode.Length / 3, 2];
        int[] etchStart = new int[gcode.Length / 3];
        int etches = -1;
        double zOldPos = 0;
        double zNew = 0;
        double[,] xyz = gcodeTOdouble(gcode);


        for (int i = 0; i < gcode.Length; i++)
        {
            string gc = gcode[i];
            if ((gc + "   ").Substring(0, 3) == "G00" || (gc + "   ").Substring(0, 3) == "G01" || (gc + "   ").Substring(0, 3) == "G0 " || (gc + "   ").Substring(0, 3) == "G1 ")
            {
                if (gc.Substring(0, 1) == "G")
                {
                    zOldPos = zNew;
                    string[] g = gc.Split(' ');
                    string gType = g[0];
                    zNew = xyz[i, 2];
                    if (zNew < 1 && zOldPos > 1)
                    {
                        etches = Math.Max(etches, 0);
                        etchStart[etches] = i - 1;
                    }
                    else if (zNew > 1 && zOldPos < 1)
                    {
                        if (etches > -1)
                        {
                            etches++;
                        }
                    }
                }
            }
        }
        Console.WriteLine("Combining etches");
        for (int i = 0; i<etches - 1; i++)
        {
	        for (int j = i+1; j<etches; j++)
	        {
		        for (int iPos = etchStart[i]+1; iPos<etchStart[i+1]; iPos++)
                {
                    double[] iXYZ = new double[3];
                    iXYZ[0] = xyz[iPos, 0];
                    iXYZ[1] = xyz[iPos, 1];
                    iXYZ[2] = xyz[iPos, 2];
                    for (int jPos = etchStart[j]+1; jPos<etchStart[j+1]; jPos++)
                    {
                        //double[] jXYZ = xyzFROMstring(gcode[jPos]);
                        double[] jXYZ = new double[3];
                        jXYZ[0] = xyz[jPos, 0];
                        jXYZ[1] = xyz[jPos, 1];
                        jXYZ[2] = xyz[jPos, 2];
                        double dX = Math.Abs(jXYZ[0] - iXYZ[0]);
				        if (dX<dist)
                        {
                            double dY = Math.Abs(jXYZ[1] - iXYZ[1]);
                            if (dY<dist)
					        {
						        if (dX == 0 || dY == 0 || dX + dY < dist || dX* dX + dY* dY < distSquare)
						        {
                                    Console.Write("   "+ i.ToString() + " [" + (etchStart[i+1] - etchStart[i]).ToString() + "] + " + j.ToString() + " [" + (etchStart[j+1] - etchStart[j]).ToString() +"]");
                                    string[] gcodeNew = new string[gcode.Length];
                                    gcode.CopyTo(gcodeNew, 0);
                                    int jLength = etchStart[j + 1] - etchStart[j];
                                    int outLine = iPos+1;
                                    for (int k = jPos; k < etchStart[j+1]; k++)
                                    {
                                        if (xyzFROMstring(gcode[k])[2] < 0.1)
                                        {
                                            gcodeNew[outLine] = gcode[k];
                                            outLine++;
                                        }
                                    }
                                    for (int k = etchStart[j]; k <= jPos; k++)
                                    {
                                        if (xyzFROMstring(gcode[k])[2] < 0.1)
                                        {
                                            gcodeNew[outLine] = gcode[k];
                                            outLine++;
                                        }
                                    }
                                    int inLine = iPos;
                                    for (inLine = iPos; inLine < etchStart[i+1]; inLine++)
                                    {
                                        gcodeNew[outLine] = gcode[inLine];
                                        outLine++;
                                    }
                                    etchStart[i+1] = outLine;
                                    for (int k = i+1; k < j; k++)
                                    {
                                        while (inLine < etchStart[k+1])
                                        {
                                           gcodeNew[outLine] = gcode[inLine];
                                            outLine++;
                                            inLine++;
                                        }
                                        etchStart[k + 1] = outLine;
                                    }
                                    for (int k = j + 1; k <= etches; k++)
                                    {
                                        etchStart[k - 1] = etchStart[k];
                                    }

                                    for (int q = 1; q < gcode.Length - 1; q++)
                                    {
                                        if (xyzFROMstring(gcodeNew[q])[2] > 0.5)
                                        {
                                            if (xyzFROMstring(gcodeNew[q - 1])[2] < 0.5 && xyzFROMstring(gcodeNew[q + 1])[2] < 0.5)
                                            {
                                                Console.WriteLine("ERROR - missing G0");
                                            }
                                        }
                                        if (gcodeNew[q].Substring(0, 2) == "G0" && q >= etchStart[0] && q < etchStart[etches-2])
                                        {
                                            bool tempFound = false;
                                            for (int r = 0; r < etches; r++)
                                            {
                                                if (etchStart[r] == q || etchStart[r] == q + 1)
                                                {
                                                    tempFound = true;
                                                }
                                            }
                                            if (!tempFound)
                                            {
                                                Console.WriteLine("Orphan G0 found");
                                            }
                                        }
                                        for (int r = 0; r < etches - 1; r++)
                                        {
                                            if (q == etchStart[r] && gcodeNew[q].Substring(0, 2) != "G0")
                                            {
                                                Console.WriteLine("ERROR - missing G0");
                                            }
                                            else if (r > 0 && q == etchStart[r] - 1 && gcodeNew[q].Substring(0, 2) != "G0")
                                            {
                                                Console.WriteLine("ERROR - prior G0 missing");
                                            }
                                            else if (r > 0 && q == etchStart[r] + 1 && gcodeNew[q].Substring(0, 2) != "G1")
                                            {
                                                Console.WriteLine("ERROR - following G1 missing");
                                            }
                                        }
                                    }


                                    etchStartEnd[etches, 0] = 0;
                                    etchStartEnd[etches, 1] = 0;
                                    etchStart[etches] = 0;
                                    gcodeNew.CopyTo(gcode, 0);
                                    etches--;
                                    j = i + 1;
                                    iPos = gcode.Length;
                                    jPos = gcode.Length;
                                    Console.WriteLine(" (" + etches.ToString() + " etches remaining)");
                                    xyz = gcodeTOdouble(gcode);
                                    success = true;
                                }
					        }
				        }
			        }
		        }
	        }
        }
        //Sort
        if (finalSort)
        {
            Console.WriteLine("Sorting");
            double[,] midPoint = new double[etches, 2];
            for (int i = 0; i < etches; i++)
            {
                midPoint[i, 0] = xyz[etchStart[i] + 1, 0];
                midPoint[i, 1] = xyz[etchStart[i] + 1, 1];
                int count = 1;
                for (int j = etchStart[i] + 2; j < etchStart[i + 1]; j++)
                {
                    midPoint[i, 0] += xyz[j, 0];
                    midPoint[i, 1] += xyz[j, 1];
                    count++;
                }
                midPoint[i, 0] /= count;
                midPoint[i, 1] /= count;
            }
            double[,] newMidpoint = midPoint;
            newMidpoint = GA_TSP.geneticPermute(newMidpoint);
            int[] order = new int[etches];
            bool orderingError = false;
            for (int i = 0; i < etches; i++)
            {
                order[i] = -1;
                for (int j = 0; j < etches; j++)
                {
                    if (Math.Abs(midPoint[j, 0] - newMidpoint[i, 0]) < 0.1 && Math.Abs(midPoint[j, 1] - newMidpoint[i, 1]) < 0.1)
                    {
                        for (int k = 0; k < i; k++)
                        {
                            if (order[k] == j) order[i] = -2;
                        }
                        if (order[i] == -1)
                        {
                            order[i] = j;
                            j = etches;
                        }
                        else order[i] = -1;
                    }
                }
                if (order[i] == -1) orderingError = true;
            }
            if (!orderingError)
            {
                string[] gcodeNew2 = new string[gcode.Length];
                int outLine2 = 0;
                while (outLine2 < etchStart[0])
                {
                    gcodeNew2[outLine2] = gcode[outLine2];
                    outLine2++;
                }
                for (int i = 0; i < etches; i++)
                {
                    for (int j = etchStart[order[i]]; j < etchStart[order[i] + 1]; j++)
                    {
                        gcodeNew2[outLine2] = gcode[j];
                        outLine2++;
                    }
                }
                gcodeNew2.CopyTo(gcode, 0);
            }
            else
            {
                Console.WriteLine("Error in ordering etches - no rearranging will be performed");
            }
        }
        if (success)
        {

            System.IO.File.Delete(gcodePath);
            using (StreamWriter objFile = File.CreateText(gcodePath))
            {
                for (int i = 0; i < gcode.Length; i++)
                {
                    objFile.WriteLine(gcode[i]);
                }
            }
        }
        return success;
        }
        

static public bool combineEtchesOld(string gcodePath, double maxLoopDistance)
    {
        int etches = 0;
        int loops = 0;
        string[] gcode = File.ReadAllLines(gcodePath, Encoding.UTF8);
        double xOldPos;
        double yOldPos;
        double zOldPos;
        double xNew = 0.0;
        double yNew = 0.0;
        double zNew = 2;
        int maxEtchLength = 0;
        int thisEtchLength = 0;
        for (int i = 0; i < gcode.Length; i++)
        {
            string gc = gcode[i];
            if ((gc + "   ").Substring(0, 3) == "G00" || (gc + "   ").Substring(0, 3) == "G01" || (gc + "   ").Substring(0, 3) == "G0 " || (gc + "   ").Substring(0, 3) == "G1 ")
            {
                if (gc.Substring(0, 1) == "G")
                {
                    zOldPos = zNew;
                    string[] g = gc.Split(' ');
                    string gType = g[0];
                    for (int j = 1; j < g.Length; j++)
                    {
                        if (g[j].Length > 0)
                        {
                            g[j] = g[j].Replace(" ", "");
                            if (g[j].Substring(0, 1) == "Z") zNew = double.Parse(g[j].Replace("Z", ""));
                        }
                    }

                    if (zNew < 1 && zOldPos > 1)
                    {
                        etches++;
                        thisEtchLength = 0;
                    }
                    else thisEtchLength++;
                    maxEtchLength = Math.Max(maxEtchLength, thisEtchLength);
                }
            }
        }
        xNew = 0.0;
        yNew = 0.0;
        zNew = 2;
        double etchSpeed = 60;
        double plungeSpeed = 30;
        double jogSpeed = 120;
        int[] etchLengths = new int[etches];
        double loopStartX = 0;
        double loopStartY = 0;
        double[,,] etchPaths = new double[gcode.Length / 5, 3, gcode.Length];
        bool[] etchIsALoop = new bool[etches];
        int etchPosition = 0;
        int thisEtch = -1;
        for (int i = 0; i < gcode.Length; i++)
        {
            string gc = gcode[i];
            if ((gc + "   ").Substring(0, 3) == "G00" || (gc + "   ").Substring(0, 3) == "G01" || (gc + "   ").Substring(0, 3) == "G0 " || (gc + "   ").Substring(0, 3) == "G1 ")
            {
                if (gc.Substring(0, 1) == "G")
                {
                    xOldPos = xNew;
                    yOldPos = yNew;
                    zOldPos = zNew;
                    gc = gc.Replace("G00", "G0");
                    gc = gc.Replace("G01", "G1");
                    gc = gc.Replace("\n", "");
                    gc = gc.Replace("\n", "");
                    string[] g = gc.Split(' ');
                    string gType = g[0];
                    for (int j = 1; j < g.Length; j++)
                    {
                        if (g[j].Length > 0)
                        {
                            g[j] = g[j].Replace(" ", "");
                            if (g[j].Substring(0, 1) == "X") xNew = double.Parse(g[j].Replace("X", ""));
                            else if (g[j].Substring(0, 1) == "Y") yNew = double.Parse(g[j].Replace("Y", ""));
                            else if (g[j].Substring(0, 1) == "Z") zNew = double.Parse(g[j].Replace("Z", ""));
                            else if (g[j].Substring(0, 1) == "F")
                            {
                                if (gType == "G0") jogSpeed = double.Parse(g[j].Replace("F", ""));
                                else if (xNew == xOldPos && yNew == yOldPos && zNew < zOldPos) plungeSpeed = double.Parse(g[j].Replace("F", ""));
                                else etchSpeed = double.Parse(g[j].Replace("F", ""));
                            }
                        }
                    }
                    if (zNew < 1)
                    {
                        if (zOldPos > 1)
                        {
                            thisEtch++;
                            loopStartX = xNew;
                            loopStartY = yNew;
                            etchLengths[thisEtch] = 0;
                            etchPosition = 0;
                            etchIsALoop[thisEtch] = false;
                        }
                        etchLengths[thisEtch]++;
                        etchPaths[thisEtch, 0, etchPosition] = xNew;
                        etchPaths[thisEtch, 1, etchPosition] = yNew;
                        etchPaths[thisEtch, 2, etchPosition] = zNew;
                        etchPosition++;
                    }
                    else if (zOldPos < 1 && Math.Abs(loopStartX - xNew) < maxLoopDistance && Math.Abs(loopStartY - yNew) < maxLoopDistance)
                    {
                        loops++;
                        etchIsALoop[thisEtch] = true;
                    }
                }
            }
        }
        etches--;
        Console.WriteLine();
        Console.Write("Total Etch / Loop Count:  ");
        Console.Write(etches);
        Console.Write(" / ");
        Console.WriteLine(loops);
        bool success = false;
        float maxLoopDistanceSquared = (float)maxLoopDistance * (float)maxLoopDistance;
        for (int i = 0; i < etches - 1; i++)
        {
            for (int j = i + 1; j < etches; j++)
            {
                if (etchIsALoop[i] == true && etchIsALoop[j] == true)
                {
                    double[,] a = new double[etchLengths[i], 3];
                    for (int n = 0; n < etchLengths[i]; n++)
                    {
                        a[n, 0] = etchPaths[i, 0, n];
                        a[n, 1] = etchPaths[i, 1, n];
                        a[n, 2] = etchPaths[i, 2, n];
                    }
                    double[,] b = new double[etchLengths[j], 3];
                    for (int n = 0; n < etchLengths[j]; n++)
                    {
                        b[n, 0] = etchPaths[j, 0, n];
                        b[n, 1] = etchPaths[j, 1, n];
                        b[n, 2] = etchPaths[j, 2, n];
                    }
                    double[,] possNewEtch = mixEtches(a, b, (float)maxLoopDistance);
                    if (possNewEtch.Length > 1)
                    {
                        Console.WriteLine("  Combining etches " + i.ToString() + " & " + j.ToString());
                        success = true;
                        for (int n = 0; n < possNewEtch.Length / 3; n++)
                        {
                            etchPaths[i, 0, n] = possNewEtch[n, 0];
                            etchPaths[i, 1, n] = possNewEtch[n, 1];
                            etchPaths[i, 2, n] = possNewEtch[n, 2];
                            etchLengths[i] = possNewEtch.Length / 3;
                        }
                        for (int n = j; n < etches; n++)
                        {
                            for (int o = 0; o < Math.Max(etchLengths[n], etchLengths[n + 1]); o++)
                            {
                                etchPaths[n, 0, o] = etchPaths[n + 1, 0, o];
                                etchPaths[n, 1, o] = etchPaths[n + 1, 1, o];
                                etchPaths[n, 2, o] = etchPaths[n + 1, 2, o];
                            }
                            etchLengths[n] = etchLengths[n + 1];
                        }
                        etches--;
                    }
                }
            }
        }
        if (!success) return false;
        using (StreamWriter objFile = File.CreateText(gcodePath.Replace(".gcode", ".tmp")))
        {
            objFile.WriteLine("M3");
            for (int I = 0; I < etches + 1; I++)
            {
                objFile.WriteLine("G0 Z1.5 F120");
                objFile.Write("G0 X");
                objFile.Write(etchPaths[I, 0, 0]);
                objFile.Write(" Y");
                objFile.Write(etchPaths[I, 1, 0]);
                objFile.Write(" F");
                objFile.WriteLine(jogSpeed);
                objFile.Write("G1 Z");
                double zLast = etchPaths[I, 2, 0];
                objFile.Write(etchPaths[I, 2, 0]);
                objFile.Write(" F");
                objFile.WriteLine(plungeSpeed);
                double feedLast = plungeSpeed;
                for (int J = 1; J <= etchLengths[I] - 1; J++)
                {
                    objFile.Write("G1 X");
                    objFile.Write(etchPaths[I, 0, J]);
                    objFile.Write(" Y");
                    objFile.Write(etchPaths[I, 1, J]);
                    if (etchPaths[I, 2, J] != zLast)
                    {
                        objFile.Write(" Z");
                        objFile.Write(etchPaths[I, 2, J]);
                        zLast = etchPaths[I, 2, J];
                    }
                    if (J == 1)
                    {
                        objFile.Write(" F");
                        objFile.Write(etchSpeed);
                    }
                    objFile.WriteLine();
                }
            }
            objFile.WriteLine("M5");
        }
        //Rename .tmp to .gcode
        //Copy data from .tmp to .gcode and add header info
        string currentContent = File.ReadAllText(gcodePath.Replace(".gcode", ".tmp"));
        if (System.IO.File.Exists(gcodePath.Replace(".tmp", ".gcode"))) File.Delete(gcodePath.Replace(".tmp", ".gcode")); //try/catch exception handling needs to be implemented

        File.WriteAllText(gcodePath, ";Etches optimized with td0g's PCB Gcode Toolkit v" + Program.currentVersion + "\n;See www.github.com/td0g \n\n" + currentContent);
        System.IO.File.Delete(gcodePath.Replace(".gcode", ".tmp"));
        return true;
    }

    static public double[,] mixEtches(double[,] a, double[,] b, float maxLoopDistance)
    {
        int l = 0;
        int m = 0;
        float maxLoopDistanceSquared = maxLoopDistance * maxLoopDistance;
        float d = 9999;
        while (l < a.Length/3)
        {
            m = 0;
            while (m < b.Length / 3)
            {
                float dX = (float)Math.Abs(a[l,0]-b[m,0]);
                if (dX < maxLoopDistance)
                {
                    float dY = (float)Math.Abs(a[l, 1] - b[m, 1]);
                    if (dY < maxLoopDistance)
                    {
                        d = dX * dX + dY * dY;
                        if (d <= maxLoopDistanceSquared)
                        {
                            //m--;
                            //l--;
                            if (m < 0) m += b.Length / 3;
                            if (l < 0) l += a.Length / 3;
                            double[,] c = new double[a.Length / 3 + b.Length / 3 + 2, 3];
                            int i = 0;
                            for (int n = 0; n <= l; n++)    //0.7.1 was n < l
                            {
                                c[i, 0] = a[n, 0];
                                c[i, 1] = a[n, 1];
                                c[i, 2] = a[n, 2];
                                i++;
                            }
                            for (int n = 0; n <= b.Length / 3; n++)   //0.7.1 was n < etchLengths
                            {
                                c[i, 0] = b[(n + m) % (b.Length / 3), 0];
                                c[i, 1] = b[(n + m) % (b.Length / 3), 1];
                                c[i, 2] = b[(n + m) % (b.Length / 3), 2];
                                i++;
                            }
                            for (int n = l; n < a.Length / 3; n++)
                            {
                                c[i, 0] = a[n, 0];
                                c[i, 1] = a[n, 1];
                                c[i, 2] = a[n, 2];
                                i++;
                            }
                            return c;
                        }
                    }
                }
                m++;
            }
            l++;
        }
        double[,] cc = new double[1, 1];
        cc[0, 0] = 0;
        return cc;
    }



    static public double[,] findHoles(string gcodePath)
    {
        double xTarget = 0;
        double yTarget = 0;
        double zTarget = 0;
        string[] gcode = File.ReadAllLines(gcodePath, Encoding.UTF8);
        double[,] hole = new double[gcode.Length, 2];
        int holeCount = 0;
        double xLast;
        double yLast;
        string gType;
        for (int i = 0; i < gcode.Length; i++)
        {
            xLast = xTarget;
            yLast = yTarget;
            string gc = gcode[i];
            if ((gc + "   ").Substring(0, 3) == "G00" || (gc + "   ").Substring(0, 3) == "G01" || (gc + "   ").Substring(0, 3) == "G0 " || (gc + "   ").Substring(0, 3) == "G1 ")
            {
                string[] g = gc.Split(' ');
                gType = g[0];
                for (int j = 1; j < g.Length; j++)
                {
                    if (g[j].Length > 0)
                    {
                        g[j] = g[j].Replace(" ", "");
                        if (g[j].Substring(0, 1) == "X") xTarget = float.Parse(g[j].Replace("X", ""));
                        else if (g[j].Substring(0, 1) == "Y") yTarget = float.Parse(g[j].Replace("Y", ""));
                        else if (g[j].Substring(0, 1) == "Z") zTarget = float.Parse(g[j].Replace("Z", ""));
                    }
                }
            }

            //########################################################################################
            //                          List all holes
            //########################################################################################


            bool duplicate = true;
            if (zTarget < 0)
            {
                duplicate = false;
                for (int j = 0; j < holeCount - 1; j++)
                {
                    if (xTarget == hole[j, 0] && yTarget == hole[j, 1])
                    {
                        duplicate = true;
                        j = holeCount;
                    }
                }
                if (!duplicate)
                {
                    holeCount = holeCount + 1;
                    hole[holeCount - 1, 0] = xTarget;
                    hole[holeCount - 1, 1] = yTarget;
                }
            }
        }
        double[,] holeOut = new double[holeCount, 2];
        for (int i = 0; i < holeCount; i++)
        {
            holeOut[i, 0] = hole[i, 0];
            holeOut[i, 1] = hole[i, 1];
        }

        return holeOut;
    }


    static public void correctZBacklash(string gcodePath, double distance)
    {
        double xyDistCritDown = 0.5;  //Maximum XY distance for gradual correction when Z is moving down 
        double xyDistCritUp = 3;    //Maximum XY distance for gradual correction when Z is moving up (should be longer than down)

        //Variables
        int gType = -1;
        double targetX = 0;
        double targetY = 0;
        double targetZ = 0;
        double[] feedRate = { 60, 30 };
        double lastX = double.NaN;
        double lastY = double.NaN;
        double lastZ = double.NaN;
        double lastOutputZ = double.NaN;
        double[] lastFeedRate = { double.NaN, double.NaN };
        int zCurrComp = 0;
        xyDistCritUp = xyDistCritUp * xyDistCritUp;
        xyDistCritDown = xyDistCritDown * xyDistCritDown;

        //Loop through gcode
        using (StreamWriter objFile = File.CreateText(gcodePath.Replace(".gcode", ".tmp")))
        {
            string[] gcode = File.ReadAllLines(gcodePath, Encoding.UTF8);
            for (int i = 0; i < gcode.Length; i++)
            {
                string gc = gcode[i];
                string[] g = gc.Split(' ');
                //Parse gcode

                if ((gc + "   ").Substring(0, 3) != "G00" && (gc + "   ").Substring(0, 3) != "G01" && (gc + "   ").Substring(0, 3) != "G0 " && (gc + "   ").Substring(0, 3) != "G1 ") objFile.WriteLine(gc);
                else if (gc.Substring(0, 1) == "G")
                {
                    gc = gc.Replace("G00", "G0");
                    gc = gc.Replace("G01", "G1");
                    gc = gc.Replace("\n", "");
                    gc = gc.Replace("\n", "");
                    gType = 1;
                    for (int j = 0; j < g.Length; j++)
                    {
                        if (g[j].Length > 0)
                        {
                            g[j] = g[j].Replace(" ", "");
                            if (g[j].Substring(0, 1) == "G")
                            {
                                gType = int.Parse(g[j].Replace("G", ""));
                            }
                            else if (g[j].Substring(0, 1) == "X")
                            {
                                targetX = double.Parse(g[j].Replace("X", ""));
                            }
                            else if (g[j].Substring(0, 1) == "Y")
                            {
                                targetY = double.Parse(g[j].Replace("Y", ""));
                            }
                            else if (g[j].Substring(0, 1) == "Z")
                            {
                                targetZ = double.Parse(g[j].Replace("Z", ""));
                            }
                            else if (g[j].Substring(0, 1) == "F")
                            {
                                feedRate[gType] = double.Parse(g[j].Replace("F", ""));
                            }
                        }
                    }
                    double correctedZ = targetZ;

                    //Are axis moving?
                    bool x = true;
                    if (Math.Round(lastX, 3) == Math.Round(targetX, 3)) x = false;
                    bool y = true;
                    if (Math.Round(lastY, 3) == Math.Round(targetY, 3)) y = false;

                    //Calc distance of XY travel
                    double xyDistSquared = Math.Pow(Math.Abs(targetX - lastX), 2) + Math.Pow(Math.Abs(targetY - lastY), 2);

                    //Do we need to compensate up or down?
                    int zThisComp = 1;
                    if (targetZ < lastZ) zThisComp = -1;

                    //If XY travel is too far for compensation direction change, then write an intermediate gcode for faster compensation
                    double xyDistCrit = xyDistCritDown;
                    if (zThisComp == 1) xyDistCrit = xyDistCritUp;
                    double outputZ;
                    if (xyDistSquared > xyDistCrit && targetZ < 0.1 && zCurrComp != zThisComp)
                    {
                        double a = Math.Sqrt(xyDistCrit) / Math.Sqrt(xyDistSquared);    //These values are squared for faster calculations above... Need to sqrt them.
                        objFile.Write("G" + gType);
                        if (x)
                        {
                            double xInt = lastX + (targetX - lastX) * a;
                            objFile.Write(" X" + Math.Round(xInt, 3).ToString());
                        }
                        if (y)
                        {
                            double yInt = lastY + (targetY - lastY) * a;
                            objFile.Write(" Y" + Math.Round(yInt, 3).ToString());
                        }
                        double zInt = lastZ + (targetZ - lastZ) * a + zThisComp * distance / 2; //Need to write Z regardless if it was supposed to move
                        outputZ = Math.Round(targetZ + zThisComp * distance / 2, 3);
                        objFile.Write(" Z" + outputZ.ToString());
                        lastOutputZ = outputZ;
                        if (lastFeedRate != feedRate) objFile.Write(" F" + feedRate[gType]);
                        objFile.WriteLine();
                        lastFeedRate[gType] = feedRate[gType];
                    }
                    zCurrComp = zThisComp;

                    //Write final gcode
                    outputZ = Math.Round(targetZ + zThisComp * distance / 2, 3);
                    bool z = true;
                    if (lastOutputZ == outputZ) z = false;
                    if (x || y || z)
                    {
                        objFile.Write("G" + gType);
                        if (x) objFile.Write(" X" + Math.Round(targetX, 3).ToString());
                        if (y) objFile.Write(" Y" + Math.Round(targetY, 3).ToString());
                        if (z) objFile.Write(" Z" + outputZ.ToString());
                        lastOutputZ = outputZ;
                        if (lastFeedRate[gType] != feedRate[gType]) objFile.Write(" F" + feedRate[gType]);
                        objFile.WriteLine();
                        lastX = targetX;
                        lastY = targetY;
                        lastZ = targetZ;
                        lastFeedRate[gType] = feedRate[gType];
                    }
                }
            }
        }
        //Rename .tmp to .gcode
        //Copy data from .tmp to .gcode and add header info
        string currentContent = File.ReadAllText(gcodePath.Replace(".gcode", ".tmp"));
        if (System.IO.File.Exists(gcodePath.Replace(".tmp", ".gcode"))) File.Delete(gcodePath.Replace(".tmp", ".gcode")); //try/catch exception handling needs to be implemented

        File.WriteAllText(gcodePath, ";Backlash compensation (" + distance.ToString() + "mm) added with td0g's PCB Gcode Toolkit v" + Program.currentVersion + "\n;See www.github.com/td0g \n\n" + currentContent);
        System.IO.File.Delete(gcodePath.Replace(".gcode", ".tmp"));
    }




    static public double showStats(string gcodePath)
    {
        double xMax = -9999999;
        double xMin = 9999999;
        double yMax = -9999999;
        double yMin = 9999999;
        double zMax = -9999999;
        double zMin = 9999999;
        double xOldPos = 0.0;
        double yOldPos = 0.0;
        double zOldPos = 0.0;
        double xNew = 0.0;
        double yNew = 0.0;
        double zNew = 0.0;
        int zeroOne = 1;
        float feedRateA = 120;
        float feedRateB = 60;
        double distThis = 0.0;
        Console.WriteLine("#####################################################################\n\n   Stats:");
        try
        {
            string[] gcode = File.ReadAllLines(gcodePath, Encoding.UTF8);
            double totalDistance = 0;
            double totalTime = 0;
            for (int i = 0; i < gcode.Length; i++)
            {
                string gc = gcode[i];
                if (gc.Length > 2)
                {
                    if ((gc + "   ").Substring(0, 3) == "G00" || (gc + "   ").Substring(0, 3) == "G01" || (gc + "   ").Substring(0, 3) == "G0 " || (gc + "   ").Substring(0, 3) == "G1 ")
                    {
                        if (gc.Substring(0, 1) == "G")
                        {
                            gc = gc.Replace("G00", "G0");
                            gc = gc.Replace("G01", "G1");
                            gc = gc.Replace("\n", "");
                            gc = gc.Replace("\n", "");
                            xOldPos = xNew;
                            yOldPos = yNew;
                            zOldPos = zNew;
                            string[] g = gc.Split(' ');
                            string gType = g[0];
                            for (int j = 1; j < g.Length; j++)
                            {
                                if (g[j].Length > 0)
                                {
                                    g[j] = g[j].Replace(" ", "");
                                    if (g[j].Substring(0, 1) == "X") xNew = float.Parse(g[j].Replace("X", ""));
                                    else if (g[j].Substring(0, 1) == "Y") yNew = float.Parse(g[j].Replace("Y", ""));
                                    else if (g[j].Substring(0, 1) == "Z") zNew = float.Parse(g[j].Replace("Z", ""));
                                    else if (g[j].Substring(0, 1) == "F")
                                    {
                                        if (gType == "G0")
                                        {
                                            feedRateA = float.Parse(g[j].Replace("F", ""));
                                            zeroOne = 0;
                                        }
                                        else
                                        {
                                            feedRateB = float.Parse(g[j].Replace("F", ""));
                                            zeroOne = 1;
                                        }
                                    }
                                }
                            }
                            if (xMax < xNew) xMax = xNew;
                            if (xMin > xNew) xMin = xNew;
                            if (yMax < yNew) yMax = yNew;
                            if (yMin > yNew) yMin = yNew;
                            if (zMax < zNew) zMax = zNew;
                            if (zMin > zNew) zMin = zNew;
                            distThis = Math.Sqrt((xNew - xOldPos) * (xNew - xOldPos) + (yNew - yOldPos) * (yNew - yOldPos) + (zNew - zOldPos) * (zNew - zOldPos));
                            totalDistance = totalDistance + distThis;
                            if (zeroOne == 0) totalTime = totalTime + distThis / feedRateA * 60;
                            else totalTime = totalTime + distThis / feedRateB * 60;
                        }
                    }
                }
            }
            Console.Write("Total Gcode Time:  ");
            Console.Write(Math.Round(totalTime, 0));
            Console.WriteLine(" seconds");
            Console.Write("Total Gcode Distance:  ");
            Console.Write(Math.Round(totalDistance, 0));
            Console.WriteLine(" mm\n");
            Console.WriteLine("           X           Y           Z");
            Console.WriteLine("Max:   " + (Math.Round(xMax, 2).ToString() + "            ").Substring(0, 12) + (Math.Round(yMax, 2).ToString() + "            ").Substring(0, 12) + Math.Round(zMax, 2).ToString());
            Console.WriteLine("Min:   " + (Math.Round(xMin, 2).ToString() + "            ").Substring(0, 12) + (Math.Round(yMin, 2).ToString() + "            ").Substring(0, 12) + Math.Round(zMin, 2).ToString());
            return totalDistance;
        }
        catch
        {
            Console.WriteLine(gcodePath + "is not a valid gcode file!");
            return 0;
        }
        
    }





    public static int counter = 0;
    public static int sectorID = 0;
    public static int sectors = 1;
    public static int permutationsDone;
    public static int permutationsTotal;
    public static int permutationsPercentage;
    public static double lastX = 0;
    public static double lastY = 0;
    public static double bestDistance = 9999999;
    public static int[] holePath = new int[1];
    public static int[] holeSectorCount = new int[1];
    public static double[,,] holeSector = new double[256, 1000, 2];
    public static int[] bestRoute = new int[1];

    static public void optimizeHoles(string gcodePath)
    {
        Console.WriteLine();
        int[] sectorStartHole = new int[1];
        int[] sectorEndHole = new int[1];
        int[] holePathNew = new int[1];
        double xMax = -9999999;
        double xMin = 9999999;
        double yMax = -9999999;
        double yMin = 9999999;
        float holeDepth = 0;

        int maxHolesPerSector = 10;
        
        

        string[] gcode = File.ReadAllLines(gcodePath, Encoding.UTF8);
        double[,] hole = findHoles(gcodePath);
        //########################################################################################
        //                          Divide Into Sectors
        //########################################################################################
        int holeCount = hole.Length /2;
        double xLast = 0;
        double yLast = 0;
        string gType = "";

        for (int i = 0; i < gcode.Length; i++)
        {
            string gc = gcode[i];
            if ((gc + "   ").Substring(0, 3) == "G00" || (gc + "   ").Substring(0, 3) == "G01" || (gc + "   ").Substring(0, 3) == "G0 " || (gc + "   ").Substring(0, 3) == "G1 ")
            {
                string[] g = gc.Split(' ');
                gType = g[0];
                for (int j = 1; j < g.Length; j++)
                {
                    if (g[j].Length > 0)
                    {
                        g[j] = g[j].Replace(" ", "");
                        if (g[j].Substring(0, 1) == "X")
                        {
                            float xTarget = float.Parse(g[j].Replace("X", ""));
                            if (xTarget > xMax) xMax = xTarget;
                            if (xTarget < xMin) xMin = xTarget;
                        }
                        else if (g[j].Substring(0, 1) == "Y")
                        {
                            float yTarget = float.Parse(g[j].Replace("Y", ""));
                            if (yTarget > yMax) yMax = yTarget;
                            if (yTarget < yMin) yMin = yTarget;
                        }
                        else if (g[j].Substring(0, 1) == "Z")
                        {
                            float zTarget = float.Parse(g[j].Replace("Z", ""));
                            holeDepth = Math.Min(holeDepth, zTarget);
                        }
                    }
                }
            }
        }
        xMax = (float)Math.Round(xMax, 2);
        xMin = (float)Math.Round(xMin, 2);
        yMax = (float)Math.Round(yMax, 2);
        yMin = (float)Math.Round(yMin, 2);

        int holeCountPerSector = holeCount;
        int xSectors = 1;
        int ySectors = 1;
        while (holeCountPerSector > maxHolesPerSector)
        {
            sectors = sectors * 2;
            if ((xMax - xMin) > (yMax - yMin))
            {
                if (xSectors > ySectors) ySectors = ySectors * 2;
                else xSectors = xSectors * 2;
            }
            else
            {
                if (ySectors > xSectors) xSectors = xSectors * 2;
                else ySectors = ySectors * 2;
            }
            holeCountPerSector = holeCount / sectors;
        }


        //########################################################################################
        //                          Build Sector Hole Tables
        //########################################################################################

        xLast = 0;
        yLast = 0;
        double xSectorMin;
        double xSectorMax;
        double ySectorDist = (yMax - yMin) / ySectors;
        double ySectorMin = ySectorDist * -1 + yMin;
        double ySectorMax = yMin;
        Array.Resize(ref holeSectorCount, sectors);

        for (int i = 0; i < xSectors; i++)
        {
            xSectorMin = (xMax - xMin) / xSectors * i + xMin;
            xSectorMax = (xMax - xMin) / xSectors * (i + 1) + xMin;
            for (int j = 0; j < ySectors; j++)
            {
                ySectorMax += ySectorDist;
                ySectorMin += ySectorDist;
                holeSectorCount[sectorID] = 0;
                Console.WriteLine("Sector: " + sectorID.ToString());
                Console.WriteLine("  X " + Math.Round(xSectorMin,1).ToString() + " -> " + Math.Round(xSectorMax,1).ToString());
                Console.WriteLine("  Y " + Math.Round(ySectorMin,1).ToString() + " -> " + Math.Round(ySectorMax,1).ToString());
                for (int k = 0; k < holeCount; k++)
                {
                    if ((xSectorMin == xMin || hole[k, 0] >= xSectorMin - 0.1) && (xSectorMax == xMax || hole[k, 0] <= xSectorMax + 0.1) && (ySectorMin == yMin || hole[k, 1] >= ySectorMin - 0.1) && (ySectorMax == yMax || hole[k, 1] <= ySectorMax + 0.1) && hole[k, 0] < 999999 && hole[k, 1] < 999999)
                    {
                        holeSector[sectorID, holeSectorCount[sectorID], 0] = hole[k, 0];
                        holeSector[sectorID, holeSectorCount[sectorID], 1] = hole[k, 1];
                        hole[k, 0] = 9999999;
                        hole[k, 1] = 9999999;
                        holeSectorCount[sectorID] = holeSectorCount[sectorID] + 1;
                    }
                }
                for (int k = 0; k < holeSectorCount[sectorID]; k++)
                {
                    xLast = holeSector[sectorID, k, 0];
                    yLast = holeSector[sectorID, k, 1];
                }
                Console.WriteLine("  Holes In Sector: " + holeSectorCount[sectorID].ToString());
                sectorID = sectorID + 1;
            }
            ySectorMax += ySectorDist;
            ySectorMin += ySectorDist;
            ySectorDist *= -1;
        }


        //########################################################################################
        //                          Get Entry / Exit Holes
        //########################################################################################


        Array.Resize(ref sectorStartHole, sectors);
        Array.Resize(ref sectorEndHole, sectors);
        Console.WriteLine();
        Console.WriteLine();

        for (int i = 0; i < sectors - 1; i++)
        {
            double bestEndDistance = 9999999;
            for (int j = 0; j < holeSectorCount[i]; j++)
            {
                for (int k = 0; k < holeSectorCount[i + 1]; k++)
                {
                    bool testThis = true;
                    if (i > 0)
                    {
                        if (j == sectorEndHole[i - 1]) testThis = false;
                    }
                    if (testThis)
                    {
                        if (bestEndDistance > (Math.Pow((holeSector[i, j, 0] - holeSector[i + 1, k, 0]), 2) + Math.Pow((holeSector[i, j, 1] - holeSector[i + 1, k, 1]), 2)))
                        {
                            bestEndDistance = (Math.Pow((holeSector[i, j, 0] - holeSector[i + 1, k, 0]), 2) + Math.Pow((holeSector[i, j, 1] - holeSector[i + 1, k, 1]), 2));
                            sectorEndHole[i] = j;
                            sectorStartHole[i + 1] = k;
                        }
                    }
                }
            }
        }

        //########################################################################################
        //                          Brute Force
        //########################################################################################

        int holesOut = 0;
        gcodePath = gcodePath.Replace(".gcode", ".tmp");
        using (StreamWriter objFile = File.CreateText(gcodePath))
        {
            objFile.WriteLine("M17");
            objFile.WriteLine("G0 Z0.9 F120");
            objFile.WriteLine("G1 Z1 F30");
            objFile.WriteLine("M3");
            for (int i = 0; i < sectors; i++)
            {
                if (holeSectorCount[i] > 0)
                {
                    Array.Resize(ref holePath, holeSectorCount[i]);
                    Array.Resize(ref holePathNew, holeSectorCount[i]);
                    Array.Resize(ref bestRoute, holeSectorCount[i]);
                    sectorID = i;
                    for (int j = 0; j < holeSectorCount[i]; j++) holePath[j] = j;
                    if (i == 0)
                    {
                        holePath[sectorEndHole[i]] = holeSectorCount[i] - 1;
                        holePath[holeSectorCount[i] - 1] = sectorEndHole[i];
                    }
                    else if (i == sectors - 1)
                    {
                        holePath[sectorStartHole[i]] = 0;
                        holePath[0] = sectorStartHole[i];
                    }
                    else
                    {
                        holePath[sectorEndHole[i]] = holeSectorCount[i] - 1;
                        holePath[holeSectorCount[i] - 1] = sectorEndHole[i];
                        for (int j = 0; j < holeSectorCount[i]; j++)
                        {
                            if (holePath[j] == sectorStartHole[i])
                            {
                                holePath[j] = holePath[0];
                                j = holeSectorCount[i];
                            }
                        }
                        holePath[0] = sectorStartHole[i];
                    }

                    Array.Resize(ref bestRoute, holeSectorCount[i]);
                    if (holeSectorCount[i] <= maxHolesPerSector && holeSectorCount[i] > 3)
                    {
                        Console.WriteLine("Beginning Brute Force on Sector " + i.ToString());
                        bestDistance = 99999999.99;
                        permutationsDone = 0;
                        permutationsTotal = 1;
                        for (int j = 1; j < holeSectorCount[i]; j++)
                        {
                            permutationsTotal *= j;
                        }
                        if (i == 0) permutationsTotal *= holeSectorCount[i];
                        permutationsPercentage = 0;
                        permutationsTotal /= 100;
                        if (i == 0) permute(0, holeSectorCount[i] - 1);
                        else if (i == sectors - 1) permute(1, holeSectorCount[i] - 2);
                        else permute(1, holeSectorCount[i] - 1);
                        //else permuteHeap(holeSectorCount[i] - 1);
                        Console.WriteLine("\r  DONE");
                        Console.WriteLine("  Permutations:" + permutationsDone.ToString());
                    }
                    //New for 0.9
                    else if (holeSectorCount[i] > maxHolesPerSector && holeSectorCount[i] < 50)
                    {
                        Console.WriteLine("Beginning Genetic Algorithm on Sector " + i.ToString());
                        Console.WriteLine("  Holes: " + holeSectorCount[i].ToString());
                        double[,] holePathForGA = new double[holeSectorCount[i], 2];
                        for (int j = 0; j < holeSectorCount[i]; j++){
                            holePathForGA[j, 0] = holeSector[i, j, 0];
                            holePathForGA[j, 1] = holeSector[i, j, 1];
                        }
                        holePathForGA = gcodeEdit.GA_TSP.geneticPermute(holePathForGA);
                        for (int j = 0; j < holeSectorCount[i]; j++){
                            bestRoute[j] = holePath[j];
                            holeSector[i, bestRoute[j], 0] = holePathForGA[j, 0];
                            holeSector[i, bestRoute[j], 1] = holePathForGA[j, 1];
                        }
                        Console.WriteLine("\r  DONE");
                    }
                    else
                    {
                        Console.WriteLine("Copying Sector " + i.ToString());
                        for (int j = 0; j < holeSectorCount[i]; j++) bestRoute[j] = holePath[j];
                        Console.WriteLine("  DONE");
                    }
                    for (int j = 0; j < holeSectorCount[i]; j++)
                    {
                        int bestRouteIndex = bestRoute[j];
                        double xNext = holeSector[i, bestRouteIndex, 0];
                        double yNext = holeSector[i, bestRouteIndex, 1];
                        holeSector[i, bestRouteIndex, 0] = 9999999;
                        holeSector[i, bestRouteIndex, 1] = 9999999;
                        objFile.WriteLine("G0 X" + (Math.Round(xNext, 2)).ToString() + " Y" + (Math.Round(yNext, 2)).ToString());
                        objFile.WriteLine("G0 Z0");
                        objFile.WriteLine("G1 Z" + Math.Round(holeDepth,2).ToString());
                        objFile.WriteLine("G0 Z1");
                        holesOut++;
                        yLast = yNext;
                        xLast = xNext;
                    }
                }
            }
            objFile.WriteLine("M5");
            objFile.WriteLine("M18");
        }

        //########################################################################################
        //                          Wrap It Up
        //########################################################################################

        //Rename .tmp to .gcode
        if (System.IO.File.Exists(gcodePath.Replace(".tmp", ".gcode"))) File.Delete(gcodePath.Replace(".tmp", ".gcode")); //try/catch exception handling needs to be implemented
        string currentContent = String.Empty;
        currentContent = File.ReadAllText(gcodePath);
        File.WriteAllText(gcodePath.Replace(".tmp", ".gcode"), ";Optimized with td0g's PCB Gcode Toolkit v" + Program.currentVersion + "\n;See www.github.com/td0g \n\n" + currentContent);
        System.IO.File.Delete(gcodePath);

        //Notify User 

        Console.WriteLine();
        Console.WriteLine("Optimized Gcode: " + gcodePath.Replace(".tmp", ".gcode"));
        Console.WriteLine();
        Console.WriteLine("Holes in original gcode: " + holeCount.ToString() + "  holes in output gcode: " + holesOut.ToString());
        if (holeCount != holesOut)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Missed These Holes:");
            for (int k = 0; k < sectors; k++)
            {
                for (int l = 0; l < holeSectorCount[k]; l++)
                {
                    if (holeSector[k, l, 0] < 9999990 && holeSector[k, l, 1] < 9999990) Console.WriteLine("     #:" + k.ToString() + "X:" + holeSector[k, l, 0].ToString() + ", Y:" + holeSector[k, l, 1].ToString());
                }
            }
        }
        int sum = 0;
        for (int i = 0; i < sectors; i++)
        {
            sum += holeSectorCount[i];
        }
        Console.ForegroundColor = ConsoleColor.White;
        Console.WriteLine("\nSectors: " + sectors.ToString() + " x:" + xSectors.ToString() + " y:" + ySectors.ToString() + "  holes:" + sum.ToString() + "  depth:" + holeDepth.ToString());
    }

    private static void permute(int l, int r)
    {
        if (l == r) permuteDistance();
        else
        {
            for (int i = l; i <= r; i++)
            {
                int temp = holePath[l];
                holePath[l] = holePath[i];
                holePath[i] = temp;
                permute(l + 1, r);
                temp = holePath[l];
                holePath[l] = holePath[i];
                holePath[i] = temp;
            }
        }
    }



    private static void permuteHeap(int r)
    {

        int z = 1;
        for (int j = 2; j < r; j++)
        {
            z *= j;
        }
        int[] c = new int[z];
        for (int j = 0; j < r; j++)
        {
            c[j] = 0;
        }
        int i = 0;
        permuteDistance();
        while (i < r)
        {
            if (c[i] < i)
            {
                if ((i % 2) == 0)
                {
                    int temp = holePath[0];
                    holePath[0] = holePath[i];
                    holePath[i] = temp;
                }
                else
                {
                    int temp = holePath[c[i]];
                    holePath[c[i]] = holePath[i];
                    holePath[i] = temp;
                }
                permuteDistance();
                c[i]++;
                i = 0;
            }
            else
            {
                c[i] = 0;
                i++;
            }
        }
    }

    private static void permuteDistance()
    {
        if (holeSectorCount[sectorID] > 5 && permutationsDone > permutationsTotal * (permutationsPercentage + 1))
        {
            permutationsPercentage++;
            Console.Write("\r" + permutationsPercentage.ToString() + " % ");
        }
        lastX = holeSector[sectorID, holePath[0], 0];
        lastY = holeSector[sectorID, holePath[0], 1];
        double distanceThisPermute = 0;
        for (int i = 1; i < holeSectorCount[sectorID]; i++)
        {
            distanceThisPermute = distanceThisPermute + Math.Pow((lastX - holeSector[sectorID, holePath[i], 0]), 2) + Math.Pow((lastY - holeSector[sectorID, holePath[i], 1]), 2);
            lastX = holeSector[sectorID, holePath[i], 0];
            lastY = holeSector[sectorID, holePath[i], 1];
            if (distanceThisPermute > bestDistance) i = holeSectorCount[sectorID];
        }
        if (distanceThisPermute < bestDistance)
        {
            bestDistance = distanceThisPermute;
            Array.Copy(holePath, bestRoute, holeSectorCount[sectorID]);
        }
        permutationsDone++;
    }














    ///////////////////////////// Travelling Salesman Problem (TSP) ///////////////////////////////

    // TSP is a famous math problem: Given a number of cities and the costs of traveling
    // from any city to any other city, what is the cheapest round-trip route that visits
    // each city exactly once and then returns to the starting city?
    //See https://www.c-sharpcorner.com/article/how-to-use-genetic-algorithm-for-traveling-salesman-problem/

    ///////////////////////////////////////////////////////////////////////////////////////////////

    public class Chromosome
    {
        public Chromosome()
        {
            // TODO: Add constructor logic here
        }

        protected double cost;
        Random randObj = new Random();
        protected int[] cityList;
        protected double mutationPercent;
        protected int crossoverPoint;

        public Chromosome(City[] cities)
        {
            bool[] taken = new bool[cities.Length];
            cityList = new int[cities.Length];
            cost = 0.0;
            for (int i = 0; i < cityList.Length; i++) taken[i] = false;
            for (int i = 0; i < cityList.Length - 1; i++)
            {
                int icandidate;
                do
                {
                    icandidate = (int)(randObj.NextDouble() * (double)cityList.Length);
                } while (taken[icandidate]);
                cityList[i] = icandidate;
                taken[icandidate] = true;
                if (i == cityList.Length - 2)
                {
                    icandidate = 0;
                    while (taken[icandidate]) icandidate++;
                    cityList[i + 1] = icandidate;
                }
            }
            calculateCost(cities);
            crossoverPoint = 1;
        }
        public void calculateCost(City[] cities)

        {
            cost = 0;
            for (int i = 0; i < cityList.Length - 1; i++)
            {
                double dist = cities[cityList[i]].proximity(cities[cityList[i + 1]]);
                cost += dist;
            }
        }
        public double getCost()
        {
            return cost;
        }

        public int getCity(int i)
        {
            return cityList[i];
        }



        public void assignCities(int[] list)
        {
            for (int i = 0; i < cityList.Length; i++) cityList[i] = list[i];
        }

        public void assignCity(int index, int value)
        {
            cityList[index] = value;
        }
        public void assignCut(int cut)
        {
            crossoverPoint = cut;
        }
        public void assignMutation(double prob)
        {
            mutationPercent = prob;
        }
        public int mate(Chromosome father, Chromosome offspring1, Chromosome offspring2)
        {
            int crossoverPostion1 = (int)((randObj.NextDouble()) * (double)(cityList.Length - crossoverPoint));
            int crossoverPostion2 = crossoverPostion1 + crossoverPoint;
            int[] offset1 = new int[cityList.Length];
            int[] offset2 = new int[cityList.Length];
            bool[] taken1 = new bool[cityList.Length];
            bool[] taken2 = new bool[cityList.Length];
            for (int i = 0; i < cityList.Length; i++)
            {
                taken1[i] = false;
                taken2[i] = false;
            }
            for (int i = 0; i < cityList.Length; i++)
            {
                if (i < crossoverPostion1 || i >= crossoverPostion2)
                {
                    offset1[i] = -1;
                    offset2[i] = -1;
                }
                else
                {
                    int imother = cityList[i];
                    int ifather = father.getCity(i);
                    offset1[i] = ifather;
                    offset2[i] = imother;
                    taken1[ifather] = true;
                    taken2[imother] = true;
                }
            }
            for (int i = 0; i < crossoverPostion1; i++)
            {
                if (offset1[i] == -1)
                {
                    for (int j = 0; j < cityList.Length; j++)
                    {
                        int imother = cityList[j];
                        if (!taken1[imother])
                        {
                            offset1[i] = imother;
                            taken1[imother] = true;
                            break;
                        }
                    }
                }
                if (offset2[i] == -1)
                {
                    for (int j = 0; j < cityList.Length; j++)
                    {
                        int ifather = father.getCity(j);
                        if (!taken2[ifather])
                        {
                            offset2[i] = ifather;
                            taken2[ifather] = true;
                            break;
                        }
                    }
                }
            }
            for (int i = cityList.Length - 1; i >= crossoverPostion2; i--)
            {
                if (offset1[i] == -1)
                {
                    for (int j = cityList.Length - 1; j >= 0; j--)
                    {
                        int imother = cityList[j];
                        if (!taken1[imother])
                        {
                            offset1[i] = imother;
                            taken1[imother] = true;
                            break;
                        }
                    }
                }
                if (offset2[i] == -1)
                {
                    for (int j = cityList.Length - 1; j >= 0; j--)
                    {
                        int ifather = father.getCity(j);
                        if (!taken2[ifather])
                        {
                            offset2[i] = ifather;
                            taken2[ifather] = true;
                            break;
                        }
                    }
                }
            }
            offspring1.assignCities(offset1);
            offspring2.assignCities(offset2);
            int mutate = 0;
            int swapPoint1 = 0;
            int swapPoint2 = 0;
            if (randObj.NextDouble() < mutationPercent)
            {
                swapPoint1 = (int)(randObj.NextDouble() * (double)(cityList.Length));
                swapPoint2 = (int)(randObj.NextDouble() * (double)cityList.Length);
                int i = offset1[swapPoint1];
                offset1[swapPoint1] = offset1[swapPoint2];
                offset1[swapPoint2] = i;
                mutate++;
            }
            if (randObj.NextDouble() < mutationPercent)
            {
                swapPoint1 = (int)(randObj.NextDouble() * (double)(cityList.Length));
                swapPoint2 = (int)(randObj.NextDouble() * (double)cityList.Length);
                int i = offset2[swapPoint1];
                offset2[swapPoint1] = offset2[swapPoint2];
                offset2[swapPoint2] = i;
                mutate++;
            }
            return mutate;
        }

        public void PrintCity(int i, City[] cities)
        {
            System.Console.WriteLine("City " + i.ToString() + ": (" + cities[cityList[i]].getx().ToString() + ", "
              + cities[cityList[i]].gety().ToString() + ")");
        }
        public int X(int i, City[] cities)
        {
            return cities[cityList[i]].getx();
        }
        public int Y(int i, City[] cities)
        {
            return cities[cityList[i]].gety();
        }
        public static void sortChromosomes(Chromosome[] chromosomes, int num)
        {
            bool swapped = true;
            Chromosome dummy;
            while (swapped)
            {
                swapped = false;
                for (int i = 0; i < num - 1; i++)
                {
                    if (chromosomes[i].getCost() > chromosomes[i + 1].getCost())
                    {
                        dummy = chromosomes[i];
                        chromosomes[i] = chromosomes[i + 1];
                        chromosomes[i + 1] = dummy;
                        swapped = true;
                    }
                }
            }
        }
    }


    public class City
    {
        public City()
        {
        }
        private int xcoord;
        private int ycoord;
        public City(int x, int y)
        {
            xcoord = x;
            ycoord = y;
        }
        public int getx()
        {
            return xcoord;
        }
        public int gety()
        {
            return ycoord;
        }
        public int proximity(City cother)
        {
            return proximity(cother.getx(), cother.gety());
        }
        public int proximity(int x, int y)
        {
            int xdiff = xcoord - x;
            int ydiff = ycoord - y;
            return (int)Math.Sqrt(xdiff * xdiff + ydiff * ydiff);
        }
    }
    public class GA_TSP
    {
        public static int cityCount;
        public static int populationSize = 5000;
        public static double mutationPercent = 0.05;
        protected int matingPopulationSize;
        protected int favoredPopulationSize;
        protected int cutLength;
        protected int generation;
        protected Thread worker = null;
        protected bool started = false;
        public static City[] cities;
        public static Chromosome[] chromosomes;
        public GA_TSP()
        {
        }
        public void setPopulationSize(int ps)
        {
            populationSize = ps;
        }
        public void setMutationPercent(double mp)
        {
            mutationPercent = mp;
        }
        public int getPopulationSize()
        {
            return populationSize;
        }
        public double getMutationPercent()
        {
            return mutationPercent;
        }
        public void Initialization()
        {
            matingPopulationSize = populationSize / 2;
            favoredPopulationSize = matingPopulationSize / 2;
            cutLength = cityCount / 5;

            // create the initial chromosomes
            chromosomes = new Chromosome[populationSize];
            for (int i = 0; i < populationSize; i++)
            {
                chromosomes[i] = new Chromosome(cities);
                chromosomes[i].assignCut(cutLength);
                chromosomes[i].assignMutation(mutationPercent);
            }
            Chromosome.sortChromosomes(chromosomes, populationSize);
            started = true;
            generation = 0;
        }

        public static double[,] geneticPermute(double[,] unsorted)
        {
            GA_TSP tsp = new GA_TSP();
            cityCount = unsorted.Length/2;
            double[,] sorted = new double[cityCount, 2];
            bool quiet = false;
            if (tsp.getPopulationSize() < cityCount) tsp.setPopulationSize(cityCount);
            cities = new City[cityCount];
            for (int i = 0; i < cityCount; i++)
            {
                cities[i] = new City((int)(unsorted[i, 0] * 100), ((int)(unsorted[i, 1] * 100)));
            }
            tsp.Initialization();
            tsp.TSPCompute(quiet);
            for (int i = 0; i < cities.Length; i++)
            {
                sorted[i, 0] = (double)chromosomes[i].X(i, cities);
                sorted[i, 0] /= 100;
                sorted[i, 1] = (double)chromosomes[i].Y(i, cities);
                sorted[i, 1] /= 100;
            }
            return sorted;
        }


        public void TSPCompute(bool q)
        {
            double thisCost = 500.0;
            double oldCost = 0.0;
            double dcost = 500.0;
            int countSame = 0;
            Random randObj = new Random();
            while (countSame < 75)     //120
            {
                generation++;
                int ioffset = matingPopulationSize;
                int mutated = 0;
                for (int i = 0; i < favoredPopulationSize; i++)
                {
                    Chromosome cmother = chromosomes[i];
                    int father = (int)(randObj.NextDouble() * (double)matingPopulationSize);
                    Chromosome cfather = chromosomes[father];
                    mutated += cmother.mate(cfather, chromosomes[ioffset], chromosomes[ioffset + 1]);
                    ioffset += 2;
                }
                for (int i = 0; i < matingPopulationSize; i++)
                {
                    chromosomes[i] = chromosomes[i + matingPopulationSize];
                    chromosomes[i].calculateCost(cities);
                }
                // Now sort the new population

                Chromosome.sortChromosomes(chromosomes, matingPopulationSize);
                double cost = chromosomes[0].getCost();
                dcost = Math.Abs(cost - thisCost);
                thisCost = cost;
                double mutationRate = 100.0 * (double)mutated / (double)matingPopulationSize;
                if (!q) System.Console.WriteLine("Generation = " + generation.ToString() + " Cost = " + thisCost.ToString() + " Mutated Rate = " + mutationRate.ToString() + "%");
                if ((int)thisCost == (int)oldCost) countSame++;
                else
                {
                    countSame = 0;
                    oldCost = thisCost;
                }
            }
        }

        public static void sortGA(string gcodePath, double oLength)
        {
            //Get hole depth
            string[] gcode = File.ReadAllLines(gcodePath, Encoding.UTF8);
            float holeDepth = 0;
            for (int i = 0; i < gcode.Length; i++)
            {
                string gc = gcode[i];
                if ((gc + "   ").Substring(0, 3) == "G00" || (gc + "   ").Substring(0, 3) == "G01" || (gc + "   ").Substring(0, 3) == "G0 " || (gc + "   ").Substring(0, 3) == "G1 ")
                {
                    string[] g = gc.Split(' ');
                    for (int j = 1; j < g.Length; j++)
                    {
                        if (g[j].Length > 0)
                        {
                            g[j] = g[j].Replace(" ", "");
                            if (g[j].Substring(0, 1) == "Z")
                            {
                                float zTarget = float.Parse(g[j].Replace("Z", ""));
                                holeDepth = Math.Min(holeDepth, zTarget);
                            }
                        }
                    }
                }
            }
            //Get holes
            double[,] holeDouble = findHoles(gcodePath);
            int holeCount = holeDouble.Length / 2;
            holeDouble = geneticPermute(holeDouble);
            int holesOut = 0;
            //Optimize
            gcodePath = gcodePath.Replace(".gcode", ".tmp");
            using (StreamWriter objFile = File.CreateText(gcodePath))
            {
                objFile.WriteLine("M17");
                objFile.WriteLine("G0 Z0.9 F120");
                objFile.WriteLine("G1 Z1 F30");
                objFile.WriteLine("M3");

                for (int i = 0; i < holeDouble.Length/2; i++)
                {
                    objFile.Write("G0 X");
                    objFile.Write(holeDouble[i, 0].ToString());
                    objFile.Write(" Y");
                    objFile.WriteLine(holeDouble[i, 1].ToString());
                    objFile.WriteLine("G0 Z0.5");
                    objFile.WriteLine("G1 Z" + holeDepth.ToString());
                    objFile.WriteLine("G1 Z1.5");
                    holesOut++;
                }
                objFile.WriteLine("M5");
                objFile.WriteLine("M18");
            }

            //########################################################################################
            //                          Wrap It Up
            //########################################################################################
            Console.WriteLine("\nDone!");
            if (gcodeEdit.showStats(gcodePath) < oLength)
            {
                if (System.IO.File.Exists(gcodePath.Replace(".tmp", ".gcode"))) File.Delete(gcodePath.Replace(".tmp", ".gcode")); //try/catch exception handling needs to be implemented
                string currentContent = String.Empty;
                currentContent = File.ReadAllText(gcodePath);
                File.WriteAllText(gcodePath.Replace(".tmp", ".gcode"), ";Optimized with td0g's PCB Gcode Toolkit v" + Program.currentVersion + "\n;See www.github.com/td0g \n\n" + currentContent);
                Console.WriteLine("\n\n  Optimized Gcode: " + gcodePath.Replace(".tmp", ".gcode"));
                Console.WriteLine("  Total Holes: " + holeCount.ToString() + "  (actual in output gcode: " + holesOut.ToString() + ")\n");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\n\n  Optimized gcode was longer - original gcode not altered");
                Console.ForegroundColor = ConsoleColor.White;
            }
            System.IO.File.Delete(gcodePath);
        }
    }



    /// <summary>
    /// Spline interpolation class.
    /// </summary>
    class SplineInterpolator
    {
        private readonly double[] _keys;
        private readonly double[] _values;
        private readonly double[] _h;
        private readonly double[] _a;

        /// <summary>
        /// Class constructor.
        /// </summary>
        /// <param name="nodes">Collection of known points for further interpolation.
        /// Should contain at least two items.</param>
        public SplineInterpolator(IDictionary<double, double> nodes)
        {
            if (nodes == null)
            {
                throw new ArgumentNullException("nodes");
            }

            var n = nodes.Count;

            if (n < 2)
            {
                throw new ArgumentException("At least two point required for interpolation.");
            }

            _keys = nodes.Keys.ToArray();
            _values = nodes.Values.ToArray();
            _a = new double[n];
            _h = new double[n];

            for (int i = 1; i < n; i++)
            {
                _h[i] = _keys[i] - _keys[i - 1];
            }

            if (n > 2)
            {
                var sub = new double[n - 1];
                var diag = new double[n - 1];
                var sup = new double[n - 1];

                for (int i = 1; i <= n - 2; i++)
                {
                    diag[i] = (_h[i] + _h[i + 1]) / 3;
                    sup[i] = _h[i + 1] / 6;
                    sub[i] = _h[i] / 6;
                    _a[i] = (_values[i + 1] - _values[i]) / _h[i + 1] - (_values[i] - _values[i - 1]) / _h[i];
                }

                SolveTridiag(sub, diag, sup, ref _a, n - 2);
            }
        }
        
        public double GetValue(double key)
        {
            int gap = 0;
            //var previous = double.MinValue;//VTG removed
            double previous = -99999;
            key += double.Epsilon;

            // At the end of this iteration, "gap" will contain the index of the interval
            // between two known values, which contains the unknown z, and "previous" will
            // contain the biggest z value among the known samples, left of the unknown z
            for (int i = 0; i < _keys.Length; i++)
            {
                if (_keys[i] <= key && _keys[i] > previous)
                {
                    previous = _keys[i];
                    gap = i + 1;
                }
            }

            var x1 = key - previous;
            var x2 = _h[gap] - x1;

            return ((-_a[gap - 1] / 6 * (x2 + _h[gap]) * x1 + _values[gap - 1]) * x2 +
                (-_a[gap] / 6 * (x1 + _h[gap]) * x2 + _values[gap]) * x1) / _h[gap];
        }


        /// <summary>
        /// Solve linear system with tridiagonal n*n matrix "a"
        /// using Gaussian elimination without pivoting.
        /// </summary>
        private static void SolveTridiag(double[] sub, double[] diag, double[] sup, ref double[] b, int n)
        {
            int i;

            for (i = 2; i <= n; i++)
            {
                sub[i] = sub[i] / diag[i - 1];
                diag[i] = diag[i] - sub[i] * sup[i - 1];
                b[i] = b[i] - sub[i] * b[i - 1];
            }

            b[n] = b[n] / diag[n];

            for (i = n - 1; i >= 1; i--)
            {
                b[i] = (b[i] - sup[i] * b[i + 1]) / diag[i];
            }
        }
    }

    public class SplineInterpolate
    {
        static int yListSizeExpanded;
        static int xListSizeExpanded;
        static double[] yListExpanded;
        static double[] xListExpanded;
        static double[,] zArrayExpanded;

        public SplineInterpolate()
        {

        }
        public static double[,] bci(double[,] xyI)
        {
            double stepSize = bicubicExpansionGridDist;
            var known = new Dictionary<double, double>();
            int i;
            for (i = 0; i < xyI.Length / 2; i++)
            {
                known.Add(xyI[i, 0], xyI[i, 1]);
            }
            i = 0;
            var scaler = new SplineInterpolator(known);
            var start = known.First().Key;
            var end = known.Last().Key;
            double[,] xyO = new double[(int)((end - start) / stepSize)+1, 2];
            i = 0;
            for (double x = start; x <= end; x += stepSize)
            {
                xyO[i,1] = scaler.GetValue(x);
                xyO[i, 0] = x;
                i++;
            }
            return xyO;
        }

        public static void expandZList(double[,] zArray, double[] xList, double[] yList, string savePath)
        {
            int xListSize = xList.Length - 1;
            int yListSize = yList.Length - 1;
            double zMax = -9999999;
            double zMin = 9999999;
            xListSizeExpanded = (int)((xList[xListSize] - xList[0]) / bicubicExpansionGridDist);
            yListSizeExpanded = (int)((yList[yListSize] - yList[0]) / bicubicExpansionGridDist);
            double[,,] xzExpanded = new double[yListSize + 1, xListSizeExpanded + 1, 2];    //Expand along X-axis
            Array.Resize<double>(ref xListExpanded, xListSizeExpanded + 1);
            Array.Resize<double>(ref yListExpanded, yListSizeExpanded + 1);
            for (int i = 0; i <= yListSize; i++)
            {
                double[,] xz = new double[xListSize + 1, 2];
                for (int j = 0; j <= xListSize; j++)
                {
                    xz[j, 0] = xList[j];
                    xz[j, 1] = zArray[j, i];
                    zMax = Math.Max(zMax, zArray[j, i]);
                    zMin = Math.Min(zMin, zArray[j, i]);
                }
                double[,] xzExpandedThis = SplineInterpolate.bci(xz);
                for (int j = 0; j <= xListSizeExpanded; j++)
                {
                    xzExpanded[i, j, 0] = xzExpandedThis[j, 0];
                    xzExpanded[i, j, 1] = xzExpandedThis[j, 1];
                    xListExpanded[j] = xzExpandedThis[j, 0];
                }
            }
            double[,] zArrayExpandedThis = new double[xListSizeExpanded + 1, yListSizeExpanded + 1];
            for (int i = 0; i <= xListSizeExpanded; i++)    //Expand along Y-axis
            {
                double[,] yz = new double[yListSize + 1, 2];
                for (int j = 0; j <= yListSize; j++)
                {
                    yz[j, 0] = yList[j];
                    yz[j, 1] = xzExpanded[j, i, 1];
                }
                double[,] yzExpandedThis = SplineInterpolate.bci(yz);
                for (int j = 0; j <= yListSizeExpanded; j++)
                {
                    zArrayExpandedThis[i, j] = yzExpandedThis[j, 1];
                    yListExpanded[j] = yzExpandedThis[j, 0];
                }
            }
            zArrayExpanded = zArrayExpandedThis;
            //Safety Check
            double maxDelta = 0;
            for (int i = 0; i < xListSizeExpanded - 2; i++)
            {
                for (int j = 0; j < yListSizeExpanded - 2; j++)
                {
                    maxDelta = Math.Max(maxDelta, Math.Round(Math.Abs(zArrayExpanded[i, j] - zArrayExpanded[i + 1, j]), 3));
                    maxDelta = Math.Max(maxDelta, Math.Round(Math.Abs(zArrayExpanded[i, j] - zArrayExpanded[i + 1, j]), 3));
                }
            }
            if (maxDelta > 0.03)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("### WARNINING ###");
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write("Max Neighbour Delta-Z: ");
                Console.WriteLine(maxDelta);
                Console.WriteLine("Extremely steep slope detected in interpolated surface.  This was likely caused by the bicubic spline algorithm.");
                Console.WriteLine("Please review output data and consider using another interpolation method");
            }

            //Output to CSV, heat map
            Console.Write("   EXPANDED Z-PROBE HEAT MAP   ");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("Highest Z Probe: ");
            Console.Write(zMax);
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.Write("   Lowest Z Probe: ");
            Console.WriteLine(zMin);
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.White;
            double printAspectRatio = 2;
            int printDimX = 100;
            int printDimY = (int)(printDimX / printAspectRatio);
            if (printDimX / printAspectRatio * (xListExpanded[xListSizeExpanded] - xListExpanded[0]) > printDimY * (yListExpanded[yListSizeExpanded] - yListExpanded[0]))
            {
                printDimX = (int)(printDimY / (yListExpanded[yListSizeExpanded] - yListExpanded[0]) * (xListExpanded[xListSizeExpanded] - xListExpanded[0]) * printAspectRatio);
            }
            else
            {
                printDimY = (int)(printDimX * (yListExpanded[yListSizeExpanded] - yListExpanded[0]) / (xListExpanded[xListSizeExpanded] - xListExpanded[0]) / printAspectRatio);
            }
            int nextPrintX = 0;
            int nextPrintY = 0;
            using (StreamWriter objFile = File.CreateText(savePath))
            {

                objFile.Write(",");
                Console.Write("       ");
                int printXPos = 0;
                for (int i = 0; i < 9; i++)
                {
                    Console.Write(((Math.Round(xListExpanded[Math.Min(xListSizeExpanded - 1, i * xListSizeExpanded / 8)], 1).ToString()) + "    ").Substring(0, 5));
                    printXPos += 5;
                    while (printXPos < printDimX * (i + 1) / 8)
                    {
                        Console.Write(" ");
                        printXPos++;
                    }
                }
                Console.WriteLine();
                Console.Write((((Math.Round(yListExpanded[yListSizeExpanded - 1], 1)).ToString()) + "     ").Substring(0, 6));
                Console.Write("  ");
                for (int i = 0; i <= xListSizeExpanded; i++)
                {
                    objFile.Write(Math.Round(xListExpanded[i], 3).ToString() + ",");
                }
                objFile.WriteLine();
                for (int i = 0; i <= yListSizeExpanded; i++)
                {
                    string yString = Math.Round(yListExpanded[yListSizeExpanded - i], 3).ToString();
                    objFile.Write(yString);
                    for (int j = 0; j <= xListSizeExpanded; j++)
                    {
                        objFile.Write(",");
                        string xString = Math.Round(zArrayExpanded[j, yListSizeExpanded - i], 3).ToString();
                        objFile.Write(xString);
                        if (j == nextPrintX && i == nextPrintY)
                        {
                            double thisZ = zArrayExpanded[j, yListSizeExpanded - i];
                            if (thisZ < (zMin + (zMax - zMin) * 0.14)) Console.ForegroundColor = ConsoleColor.Black;
                            else if (thisZ < (zMin + (zMax - zMin) * 0.29)) Console.ForegroundColor = ConsoleColor.DarkBlue;
                            else if (thisZ < (zMin + (zMax - zMin) * 0.42)) Console.ForegroundColor = ConsoleColor.Blue;
                            else if (thisZ < (zMin + (zMax - zMin) * 0.57)) Console.ForegroundColor = ConsoleColor.White;
                            else if (thisZ < (zMin + (zMax - zMin) * 0.71)) Console.ForegroundColor = ConsoleColor.Red;
                            else if (thisZ < (zMin + (zMax - zMin) * 0.85)) Console.ForegroundColor = ConsoleColor.DarkRed;
                            else Console.ForegroundColor = ConsoleColor.Black;
                            Console.Write("#");
                            nextPrintX = nextPrintX + Math.Max((xListSizeExpanded / printDimX),1);
                        }

                    }
                    if (i >= nextPrintY)
                    {
                        nextPrintY += Math.Max((yListSizeExpanded / printDimY), 1);
                        if (nextPrintY <= yListSizeExpanded)
                        {
                            Console.WriteLine();
                            Console.ForegroundColor = ConsoleColor.White;
                            Console.Write((((Math.Round(yListExpanded[yListSizeExpanded - nextPrintY], 1)).ToString()) + "     ").Substring(0, 6));
                            Console.Write("  ");
                        }
                    }
                    nextPrintX = 0;
                    objFile.WriteLine();
                }
            }
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("\n\n");
            Console.WriteLine("Bicubic Interpolation Expanded Array file: " + (savePath));

        }


        static public double getNearest(double x, double y)
        {
            if (x < xListExpanded[0])
            {
                double dX = x - xListExpanded[0];
                int nearestY = 0;
                while (y > yListExpanded[nearestY] && nearestY < yListSizeExpanded) nearestY++;
                double slopeY = (zArrayExpanded[0, nearestY] - zArrayExpanded[1, nearestY]);
                slopeY /= (xListExpanded[0] - xListExpanded[1]);
                double z1 = zArrayExpanded[nearestY, 0] + dX * slopeY;
                if (y < yListExpanded[0])
                {
                    double dY = y - yListExpanded[0];
                    double slopeX = (zArrayExpanded[0, 0] - zArrayExpanded[0, 1]);
                    slopeX /= (yListExpanded[0] - yListExpanded[1]);
                    return (z1 * Math.Abs(dX) + (zArrayExpanded[0, 0] + dY * slopeX * Math.Abs(dY))) / (Math.Abs(dX) + Math.Abs(dY));

                }
                else if (y > yListExpanded[yListSizeExpanded])
                {
                    double dY = y - yListExpanded[yListSizeExpanded];
                    double slopeX = (zArrayExpanded[0, yListSizeExpanded] - zArrayExpanded[0, yListSizeExpanded - 1]);
                    slopeX /= (yListExpanded[yListSizeExpanded] - yListExpanded[yListSizeExpanded - 1]);
                    return (z1 * Math.Abs(dX) + (zArrayExpanded[0, yListSizeExpanded] + dY * slopeX * Math.Abs(dY))) / (Math.Abs(dX) + Math.Abs(dY));
                }
                else
                {
                    return z1;
                }
            }
            else if (x > xListExpanded[xListSizeExpanded])
            {
                double dX = x - xListExpanded[xListSizeExpanded];
                int nearestY = 0;
                while (y > yListExpanded[nearestY] && nearestY < yListSizeExpanded) nearestY++;
                double slopeY = (zArrayExpanded[xListSizeExpanded, nearestY] - zArrayExpanded[xListSizeExpanded - 1, nearestY]);
                slopeY /= (yListExpanded[yListSizeExpanded] - yListExpanded[yListSizeExpanded - 1]);
                double z1 = zArrayExpanded[xListSizeExpanded, nearestY] + dX * slopeY;
                if (y < yListExpanded[0])
                {
                    double dY = y - yListExpanded[0];
                    double slopeX = (zArrayExpanded[xListSizeExpanded, 0] - zArrayExpanded[xListSizeExpanded, 1]);
                    slopeX /= (yListExpanded[0] - yListExpanded[1]);
                    double z2 = zArrayExpanded[xListSizeExpanded, 0] + dY * slopeX;
                    double o = (z1 * Math.Abs(dX) + (z2 * Math.Abs(dY))) / (Math.Abs(dX) + Math.Abs(dY));
                    return o;

                }
                else if (y > yListExpanded[yListSizeExpanded])
                {
                    double dY = y - yListExpanded[yListSizeExpanded];
                    double slopeX = (zArrayExpanded[xListSizeExpanded, yListSizeExpanded] - zArrayExpanded[xListSizeExpanded, yListSizeExpanded - 1]);
                    slopeX /= (yListExpanded[yListSizeExpanded] - yListExpanded[yListSizeExpanded - 1]);
                    return (z1 * Math.Abs(dX) + (zArrayExpanded[xListSizeExpanded, yListSizeExpanded] + dY * slopeX * Math.Abs(dY))) / (Math.Abs(dX) + Math.Abs(dY));
                }
                else
                {
                    return z1;
                }
            }
            else if (y < yListExpanded[0])
            {
                double dY = y - yListExpanded[0];
                int nearestX = 0;
                while (x > xListExpanded[nearestX] && nearestX < xListSizeExpanded) nearestX++;
                double slopeX = (zArrayExpanded[nearestX, 0] - zArrayExpanded[nearestX, 1]);
                slopeX /= (yListExpanded[0] - yListExpanded[1]);
                return zArrayExpanded[nearestX, 0] + dY * slopeX;

            }
            else if (y > yListExpanded[yListSizeExpanded])
            {
                double dY = y - yListExpanded[yListSizeExpanded];
                int nearestX = 0;
                while (x > xListExpanded[nearestX]) nearestX++;
                double slopeX = (zArrayExpanded[nearestX, yListSizeExpanded] - zArrayExpanded[nearestX, yListSizeExpanded - 1]);
                slopeX /= (yListExpanded[yListSizeExpanded] - yListExpanded[yListSizeExpanded - 1]);
                return zArrayExpanded[nearestX, yListSizeExpanded] + dY * slopeX;
            }
            else
            {
                int nearestX = 0;
                if (x > xListExpanded[0])
                {
                    if (x > xListExpanded[xListSizeExpanded-1]) nearestX = xListSizeExpanded;
                    else
                    {
                        while (x > xListExpanded[nearestX]) nearestX++;
                        if (Math.Abs(x - xListExpanded[nearestX - 1]) < Math.Abs(x - xListExpanded[nearestX])) nearestX--;
                    }
                }
                int nearestY = 0;
                if (y > yListExpanded[0])
                {
                    if (y > yListExpanded[yListSizeExpanded]) nearestY = yListSizeExpanded;
                    else
                    {
                        while (y > yListExpanded[nearestY]) nearestY++;
                        if (Math.Abs(y - yListExpanded[nearestY - 1]) < Math.Abs(y - yListExpanded[nearestY])) nearestY--;
                    }
                }
                return zArrayExpanded[nearestX, nearestY];
            }
        }
    }
}


public class Form1 : Form
{
    public Button button1;
    public static double[,] toolPath = new double[100000,3];
    public static int toolPathLength = 0;
    public static double xMax = -9999999;
    public static double xMin = 9999999;
    public static double yMax = -9999999;
    public static double yMin = 9999999;
    public static double zMaxTravel = -9999999;
    public static double zMax = -9999999;
    public static double zMin = 9999999;
    public static double margin = 50;
    public static int maxH = 800;
    public static int maxW = 1200;
    public static double scale = 1;
    public static bool shutDown = false;

    public Form1()
    {
        //button1 = new Button();
        //button1.Size = new Size(40, 40);
        //button1.Location = new Point(30, 30);
        //button1.Text = "Click me";
        double xSize = (xMax - xMin);
        double ySize = (yMax - yMin);
        if (xSize / maxW > ySize / maxH) scale = maxW / xSize;
        else scale = maxH / ySize;
        this.Text = "Gcode Path";
        this.Width = (int)(xSize * scale + margin*2);
        this.Height = (int)(ySize * scale + margin*2);
        this.RtlTranslateContent(ContentAlignment.TopLeft);

        //this.Controls.Add(button1);

        //button1.Click += new EventHandler(button1_Click);
        if (shutDown)
        {
            this.Close();
        }

    }
    protected override void OnPaint(PaintEventArgs pea)
    {
        // Defines pen 
        Pen pen = new Pen(ForeColor,1);
        Pen penTravel = new Pen(Color.Green,1);
        for (int i = 1; i < toolPathLength; i++)
        {
            float x0 = (float)((toolPath[i - 1, 0] - xMin) * scale + margin);
            float y0 = (float)(((yMax - yMin)  - toolPath[i - 1, 1] + yMin) * scale + margin);
            float x1 = (float)((toolPath[i, 0] - xMin)* scale + margin);
            float y1 = (float)(((yMax - yMin) - toolPath[i, 1] + yMin)* scale + margin);
            if (toolPath[i - 1, 2] > 0.5 && toolPath[i, 2] > 0.5) pea.Graphics.DrawLine(penTravel, x0, y0, x1, y1);
            else if (toolPath[i - 1, 2] < 0.5 && toolPath[i, 2] < 0.5)
            {
                double col = 0.5;
                if (zMax != zMin) col = (((toolPath[i - 1, 2] + toolPath[i, 2]) / 2) - zMin)/(zMax - zMin);
                pen.Color = Color.FromArgb(255, (int)(255 * col), 0, (int)(255 * (double)(1 - col)));
                pea.Graphics.DrawLine(pen, x0, y0, x1, y1);
            }
        }
    }
    private void button1_Click(object sender, EventArgs e)
    {
        MessageBox.Show("Hello World");
        
    }
    [STAThread]
    public static void runForm(string gcodePath)
    {
        string[] gcode = File.ReadAllLines(gcodePath, Encoding.UTF8);
        double xOldPos = 0.0;
        double yOldPos = 0.0;
        double zOldPos = 0.0;
        double xNew = 0.0;
        double yNew = 0.0;
        double zNew = 0.0;
        toolPathLength = 0;
        for (int i = 0; i < gcode.Length; i++)
        {
            string gc = gcode[i];
            if (gc.Length > 2)
            {
                if ((gc + "   ").Substring(0, 3) == "G00" || (gc + "   ").Substring(0, 3) == "G01" || (gc + "   ").Substring(0, 3) == "G0 " || (gc + "   ").Substring(0, 3) == "G1 ")
                {
                    if (gc.Substring(0, 1) == "G")
                    {
                        gc = gc.Replace("G00", "G0");
                        gc = gc.Replace("G01", "G1");
                        gc = gc.Replace("\n", "");
                        gc = gc.Replace("\n", "");
                        xOldPos = xNew;
                        yOldPos = yNew;
                        zOldPos = zNew;
                        string[] g = gc.Split(' ');
                        string gType = g[0];
                        for (int j = 1; j < g.Length; j++)
                        {
                            if (g[j].Length > 0)
                            {
                                g[j] = g[j].Replace(" ", "");
                                if (g[j].Substring(0, 1) == "X") xNew = float.Parse(g[j].Replace("X", ""));
                                else if (g[j].Substring(0, 1) == "Y") yNew = float.Parse(g[j].Replace("Y", ""));
                                else if (g[j].Substring(0, 1) == "Z") zNew = float.Parse(g[j].Replace("Z", ""));
                            }
                        }
                        if (xMax < xNew) xMax = xNew;
                        if (xMin > xNew) xMin = xNew;
                        if (yMax < yNew) yMax = yNew;
                        if (yMin > yNew) yMin = yNew;
                        if (zMaxTravel < zNew) zMaxTravel = zNew;
                        if (zMax < zNew && zNew < 0.5) zMax = zNew;
                        if (zMin > zNew)
                        {
                            zMin = zNew;
                        }
                        toolPath[toolPathLength, 0] = xNew;
                        toolPath[toolPathLength, 1] = yNew;
                        toolPath[toolPathLength, 2] = zNew;
                        toolPathLength++;
                    }
                }
            }
        }


        Application.EnableVisualStyles();
        Thread _thread = new Thread(() =>{Application.Run(new Form1());});
        //Form form = new Form1();
        //Thread _thread = new Thread(new ThreadStart(new Form1()));
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
        
    }

    public static void shutDownForms()
    {
        shutDown = true;
    }
}

