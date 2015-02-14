# Num2Word
Utility to convert a number to words using the Hindi numeral system.

#Build Instructions
You will need Visual Studio 2013 (Community Edition will suffice) to build this project.
You can also build this using the .NET Framework SDK, although the build steps would be a bit more involved.
The binary files can be used directly from the "Debug"/"Release" directories of the Num2Word C# project.

#Usage Instructions
You need Excel 2003 or later to use this plugin.
1. Double click the file "Num2Word-AddIn.xll" in the output folder (Debug/Release). This will open up Excel on your PC.

2. Excel might show a warning regarding unavailability of digital signature. Ignore this by clicking on "Enable add-in for this session only".

3. Create new workbook or open any previously saved workbook you want to work on.

4. You can convert any number (whole numbers only) to its Hindi word-equivalent by typing N2W(<cell reference>)  as an excel formula in any cell.
