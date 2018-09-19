Attribute VB_Name = "mdlBuka"
Option Explicit

Public dbMesin As Database
Public tblMacam As Recordset
Public tblJenis As Recordset
Public tblCiri As Recordset
Public tblPasswd As Recordset
Public tblRelasi1 As Recordset
Public tblRelasi2 As Recordset

Public Sub Buka()
    Set dbMesin = DBEngine.Workspaces(0).OpenDatabase(App.Path + "\dbMesin.mdb")
    Set tblMacam = dbMesin.OpenRecordset("tblMacam", dbOpenTable)
    Set tblJenis = dbMesin.OpenRecordset("tblJenis", dbOpenTable)
    Set tblCiri = dbMesin.OpenRecordset("tblCiri", dbOpenTable)
    Set tblPasswd = dbMesin.OpenRecordset("tblPasswd", dbOpenTable)
    Set tblRelasi1 = dbMesin.OpenRecordset("tblRelasi1", dbOpenTable)
    Set tblRelasi2 = dbMesin.OpenRecordset("tblRelasi2", dbOpenTable)
End Sub
