# ACPI-Tools-And-Frames

ACPI_Registry_Tables_v01.xlsb
==========================================================================================================================
PURPOSE
WinOS ACPI (Advanced Configuration and Power Interface) table extraction from registry. 

GOAL
Utilize VBA and Windows DLL functions to access/parse registry keys and values. Display ACPI table values for inspection and interpret the simpler ones.

PROCESSING NOTE
The following code resides in Sub subDisplayACPI_Table(sWSName As String):
      If (lUBound > 131) Then lUBound = 500  
It allows for debug style processing by limiting any table size to 500 bytes.  Comment this line out to process until max table size. See note below about DoEvents interrupts.

HIGH LEVEL SPEC
Uses VBA to access the registry. On MAIN tab set up Registry Key Root as HKEY_LOCAL_MACHINE, Registry Key Start as HARDWARE\ACPI. The Output File Name is set as ACPI_registry.txt, but this be changed as needed (this file is a stub, is not being populated, but is available if so desired).  Click READ ACPI button to execute code in Module 1 (access this via Developer->Visual Basic interface). All subs and functions reside in Module 1.

Tool utilizes Windows dll functions to access the registry. It traverses and recursively parses registry keys until the REG_BINARY level is encountered. It pulls that data into a dynamic array of bytes.  See code in subParseSubKeys(sTableName As String, sKey As String) subroutine:
    ReDim resBinary(0 To lDataLen - 1) As Byte
    rVal = RegEnumValue(hTempKey, 0, sName, lNameLen, ByVal 0&, lValueType, resBinary(0), lDataLen)

A tab for the particular registry subkey corresponding to an ACPI table (e.g. FADT) is created (if tab already exists, it will be deleted and recreated new). ACPI table is then listed in hex and unicode for visual inspection with simple formatting. The bytes size of the table will be displayed in top corner on each tab.

OUTPUT
The MAIN tab displays the summary of what is processed. It is a mirror image of what you find in the actual Windows registry for HARDWARE\ACPI.

FORMATTING
There is a lot of formatting going on to output data in a useful form. If a simpler style output is needed, leverage the text output file.  Additionally, FADT and RSDT have been fully interpreted to show their Header values as well as their contents. Tab TableFields are the exact ACPI specification for these tables and shows the field layout breakdown. It was coded as such.

STATUS BAR
Status bar shows processing activity. Because code uses DoEvents as interrupt to allow other CPU activities, the time it takes to complete all work will vary with the size of the ACPI tables and your particular machine.  

OTHER TABS and INFO
Other tabs are present and contain support and documentation info. 

SOURCES:
-  http://www.acpi.info/
-  https://www.acpica.org/
-  https://docs.microsoft.com/en-us/windows/desktop/SysInfo/registry-value-types
-  stackoverflow.com
