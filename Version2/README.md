#Client-Dispatch-Version2

This version of client-dispatch has the ability to copy ChemStructures between Excel speadsheets by using VBA. Python code links to VBA macros via win32com.client and passes to the macro-enabled Excel file "Dispatch_control" required arguments such as target sheet and source sheet file paths for execution of the structure_copy macro. However, this only works on Windows and even then there are bugs which prevent the macro executing in the desired manner (the macros seems to start executing before the Excel file has updated to the latest version of ChemOffic for example). Version 3 will by-pass this issue by providing a master dispatch template that has a copy_structures macro already built in and hence not dependent on COM objects. 