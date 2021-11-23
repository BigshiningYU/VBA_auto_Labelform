# EXCEL_Mergingfiles
Using EXCEL VBA, you can merge the files


Sub SearchFile()

 
   
'taehee book  make a file where you will merge the files
    
     Dim strwkbook As String
     Dim wk As Variant
     
     strwkbook = ThisWorkbook.Path
     ChDir (strwkbook)
     Set wk = Workbooks.Add
     wk.SaveAs Filename:=strwkbook & "\" & "taehee book.xlsx"

      
'Open the dialog box and choose the workbook files to merge

    Dim fd As FileDialog
    Dim pathSelectedItem As Variant
    
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
   
    fd.AllowMultiSelect = True
   
    If fd.Show = -1 Then
   
    For Each pathSelectedItem In fd.SelectedItems

'Open the workbook files which you have choosed to merge

     Workbooks.Open Filename:=pathSelectedItem
   
     Cells.Select
     Selection.Copy
   
     Workbooks.Open Filename:="taehee book.xlsx" 
     Sheets.Add After:=ActiveSheet
     ActiveSheet.Paste
     
'Change the name of the sheets

     Dim fname As String
     fname = Mid(pathSelectedItem, InStrRev(pathSelectedItem, "\") + 1)
     ActiveSheet.Name = fname
   
'Close the excel file

     Workbooks(fname).Close
     
    Next pathSelectedItem
    
    Else
    
    End If

   
    
End Sub
