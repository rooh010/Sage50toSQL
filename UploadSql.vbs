Option Compare Database


Public Sub uploadSql1()

Dim sTblNm As String
Dim sTypExprt As String
Dim sCnxnStr As String, vStTime As Variant
Dim db As Database, tbldef As DAO.TableDef


sTypExprt = "ODBC Database"
sCnxnStr = "ODBC;DSN=sagesql;UID=sa;PWD=password"
vStTime = Timer
Application.Echo False, "Visual Basic code is executing."

Set db = CurrentDb()

For Each tbldef In db.TableDefs
Debug.Print tbldef.Name
sTblNm = tbldef.Name
DoCmd.TransferDatabase acExport, sTypExprt, sCnxnStr, acTable, sTblNm, sTblNm
Next tbldef


SmoothExit_ExportTbls:
Set db = Nothing
Application.Echo True
DoCmd.Quit

Exit Sub

DoCmd.Quit

End Sub



Public Function runUpload()

Call uploadSql1

End Function