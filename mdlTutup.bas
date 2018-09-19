Attribute VB_Name = "mdlTutup"
Option Explicit

Public Sub Tutup()
    tblMacam.Close
    tblJenis.Close
    tblCiri.Close
    tblPasswd.Close
    tblRelasi1.Close
    tblRelasi2.Close
    dbMesin.Close
End Sub
