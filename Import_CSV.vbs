
const pagStaff = 1

camino = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName) - (Len(WScript.ScriptName) + 1)))
Set objExcel = CreateObject("Excel.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWorkbook1 = objExcel.Workbooks.Open(camino&"\Registro_Imagenes.xlsx")
Set objWorksheet1 = objWorkbook1.Sheets(pagStaff)
objWorksheet1.Activate

objWorksheet1.Range("I:J").EntireColumn.Delete
objWorksheet1.Range("A:G").EntireColumn.Delete


If objFSO.FileExists(camino&"\"&"imgsUrls.csv") Then

objFSO.DeleteFile camino&"\"&"imgsUrls.csv",True
End If

objWorkbook1.SaveAs camino&"\"&"imgsUrls.csv",6, False, True
MsgBox "Conversion Realizada"

objWorkbook1.close False