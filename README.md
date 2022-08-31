# PolyPeptitePurityOptimizer
It is to calculate the optimized mixing portions of polypeptide with different levels of impurities.
Installation Steps:

1.	Install python 2.7 in your computer. If the current system is 64 bit, use the installer in this folder. If the current system is 32 bit, google for python 2.7 for 32 bit systems.
2.	When installing python, make sure the destination directory is at "C:\Python27\"
3.	Install "PuLP" package
a.	Open a command prompt by clicking “Run” in the Start Menu, and type ‘cmd’ in the window and push enter.
b.	Navigate to the extracted folder with the setup file in it. [Do this by typing ‘cd foldername’ at the prompt, where ‘cd’ stands for current directory and the ‘foldername’ is the name of the folder to open in the path already listed to the left of the prompt. To return back to a root drive, type ‘cd C:’]
c.	Type ‘setup.py install’ at the command prompt. This will install all the PuLP functions into Python’s site-packages directory.

4.	Install "xrld" package using the same method.
5.	Open “data.xlsm” excel file. Enable Editing and Enable Macro Usage. 
Enable Macro: Excel-> “File” tab -> “Options” -> “Trust Center” ->“Trust Center Settings” -> “Macro Settings” -> Enable all Macro Settings”
6.	Edit the “Entry” Tab. The area that could be edited is the colored areas.
a.	Enter the Standard Information
b.	Enter the Peptide weight and Information from each run (Total of 6)
c.	Edit the Internal Spec when Necessary. 
1)	First try to set the purity to min: 91.5% and max: 91.5%. If the calculated yield is too low, gradually widen the range to USP specification.
2)	Each impurity limit could be lower in order to control the quality of the lot.
d.	Click the “Yield Optimization” button to maximize the yield. The solver will figure out the amount to be pulled from each bottle, which is shown in “Solution Weight (g)” column of each bottle. If the amount is different from the quantity provided, the excel cell will turn red.
e.	There are charts underneath to indicate the positions of the optimized result corresponding to the limits. You can use them to tell the limiting factors and further optimize the result.
f.	If it is necessary to change the format of the files, use the department name of the author of this file (3 letters).
