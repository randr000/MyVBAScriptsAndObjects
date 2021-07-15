# MyVBAScriptsAndObjects

**Pre_And_Post_Run.bas**

Meant to be used in an Excel file.

Pre_Run prevents the screen from updating when a macro is run and turns off automatic calculation. This makes a macro run faster. It also turns off display alerts so no pop ups asking the user to confirm something pop up. Such as when deleting a sheet.

Post_Run reverses the values that were set in Pre_Run. Probably not necessary except for updating Application.Calculation but better to be explicit. 

**clsSet.cls**

A class that creates a wrapper for dictionary objects so that they function as sets.
