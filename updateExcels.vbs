Set fso = CreateObject("Scripting.FileSystemObject")
Set xl  = CreateObject("Excel.Application")
xl.Visible = False


'                              SET FOLDER HERE
'                                  #####
'                                  #####
'                                #########
'                                 #######
'                                  #####
'                                   ###
'                                    #
For Each f In fso.GetFolder("C:\WIP\ExcelAutoUpdate").Files
  If LCase(fso.GetExtensionName(f.Name)) = "xlsx" Then
    Set wb = xl.Workbooks.Open(f.Path)
    wb.RefreshAll
    wb.Save
    wb.Close
  End If
Next

xl.Quit
