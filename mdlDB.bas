Attribute VB_Name = "mdlDB"
Option Explicit

'mendifinisikan database
Dim dbMesin As Database

'Tables
Dim tblMacam As TableDef
Dim tblJenis As TableDef
Dim tblCiri As TableDef
Dim tblRelasi1 As TableDef
Dim tblRelasi2 As TableDef
Dim tblPasswd As TableDef

'Fields
Dim NoMacam As Field
Dim Macam As Field
Dim NoJenis As Field
Dim Jenis As Field
Dim Gejala As Field
Dim NoCiri As Field
Dim Ciri As Field
Dim Diagnosa As Field
Dim Nama As Field
Dim Passwd As Field

'Index
Dim idMacam As index
Dim idJenis As index
Dim idCiri As index

'Relasi
Dim Relasi1 As Relation
Dim Relasi2 As Relation
Dim Relasi3 As Relation
Dim Relasi4 As Relation

Sub buatDB() 'Aku harap ini benar
    Set dbMesin = DBEngine.Workspaces(0).CreateDatabase(App.Path + "\dbMesin.mdb", dbLangGeneral)
    'Membuat Table MacamKerusakan
    Set tblMacam = dbMesin.CreateTableDef("tblMacam")
    Set NoMacam = tblMacam.CreateField("NoMacam", dbText, 4)
    Set Macam = tblMacam.CreateField("Macam", dbText, 75)
    'Menambahkan Field
    With tblMacam
        .Fields.Append NoMacam
        .Fields.Append Macam
    End With
    'Index tblMacam
    Set idMacam = tblMacam.CreateIndex("idMacam")
    Set NoMacam = idMacam.CreateField("NoMacam")
    idMacam.Primary = True
    idMacam.Unique = True
    idMacam.Required = True
    idMacam.Fields.Append NoMacam
    tblMacam.Indexes.Append idMacam
    
    'Membuat table jenis kerusakan
    Set tblJenis = dbMesin.CreateTableDef("tblJenis")
    Set NoJenis = tblJenis.CreateField("NoJenis", dbText, 4)
    Set Jenis = tblJenis.CreateField("Jenis", dbText, 75)
    Set Gejala = tblJenis.CreateField("Gejala", dbMemo)
    'Menambahkan Fields
    With tblJenis
        .Fields.Append NoJenis
        .Fields.Append Jenis
        .Fields.Append Gejala
    End With
    'Indexs
    Set idJenis = tblJenis.CreateIndex("idJenis")
    Set NoJenis = idJenis.CreateField("NoJenis")
    idJenis.Primary = True
    idJenis.Unique = True
    idJenis.Required = True
    idJenis.Fields.Append NoJenis
    tblJenis.Indexes.Append idJenis
    
    'Membuat Table ciri Kerusakan
    Set tblCiri = dbMesin.CreateTableDef("tblCiri")
    Set NoCiri = tblCiri.CreateField("NoCiri", dbText, 4)
    Set Ciri = tblCiri.CreateField("Ciri", dbText, 75)
    Set Diagnosa = tblCiri.CreateField("Diagnosa", dbMemo)
    'Menambahkan field
    With tblCiri
        .Fields.Append NoCiri
        .Fields.Append Ciri
        .Fields.Append Diagnosa
    End With
    'Index
    Set idCiri = tblCiri.CreateIndex("idCiri")
    Set NoCiri = idCiri.CreateField("NoCiri")
    idCiri.Primary = True
    idCiri.Unique = True
    idCiri.Required = True
    idCiri.Fields.Append NoCiri
    tblCiri.Indexes.Append idCiri
    
    'Membuat Relasi1
    Set tblRelasi1 = dbMesin.CreateTableDef("tblRelasi1")
    Set NoMacam = tblRelasi1.CreateField("NoMacam", dbText, 4)
    Set NoJenis = tblRelasi1.CreateField("NoJenis", dbText, 4)
    ' Menambahkan field ke dalam table
    With tblRelasi1
        .Fields.Append NoMacam
        .Fields.Append NoJenis
    End With
    'Membuat indexs untuk table relasi 1
    Set idMacam = tblRelasi1.CreateIndex("idMacam")
    Set NoMacam = idMacam.CreateField("NoMacam")
    idMacam.Primary = False
    idMacam.Unique = False
    idMacam.Required = False
    idMacam.Fields.Append NoMacam
    tblRelasi1.Indexes.Append idMacam
    Set idJenis = tblRelasi1.CreateIndex("idJenis")
    Set NoJenis = idJenis.CreateField("NoJenis")
    idJenis.Primary = False
    idJenis.Unique = False
    idJenis.Required = False
    idJenis.Fields.Append NoJenis
    tblRelasi1.Indexes.Append idJenis
    
    'Membuat Relasi2
    Set tblRelasi2 = dbMesin.CreateTableDef("tblRelasi2")
    Set NoJenis = tblRelasi2.CreateField("NoJenis", dbText, 4)
    Set NoCiri = tblRelasi2.CreateField("NoCiri", dbText, 4)
    'Menambahkan field ke table
    With tblRelasi2
        .Fields.Append NoJenis
        .Fields.Append NoCiri
    End With
    'Membuat indexs untuk table relasi2
    Set idJenis = tblRelasi2.CreateIndex("idJenis")
    Set NoJenis = idJenis.CreateField("NoJenis")
    With idJenis
        .Primary = False
        .Unique = False
        .Required = False
        .Fields.Append NoJenis
    End With
    tblRelasi2.Indexes.Append idJenis
    Set idCiri = tblRelasi2.CreateIndex("idCiri")
    Set NoCiri = idCiri.CreateField("NoCiri")
    With idCiri
        .Primary = False
        .Unique = False
        .Required = False
        .Fields.Append NoCiri
    End With
    tblRelasi2.Indexes.Append idCiri
    
    Set tblPasswd = dbMesin.CreateTableDef("tblPasswd")
    Set Nama = tblPasswd.CreateField("Nama", dbText, 10)
    Set Passwd = tblPasswd.CreateField("Passwd", dbText, 8)
    tblPasswd.Fields.Append Nama
    tblPasswd.Fields.Append Passwd
    dbMesin.TableDefs.Append tblPasswd
    
    'Menyimpan Table
    With dbMesin
        .TableDefs.Append tblMacam
        .TableDefs.Append tblJenis
        .TableDefs.Append tblCiri
        .TableDefs.Append tblRelasi1
        .TableDefs.Append tblRelasi2
    End With
    
    'Membuat relasi table Macam dan Relasi1
    Set Relasi1 = dbMesin.CreateRelation("Relasi1", "tblMacam", "tblRelasi1", dbRelationLeft + dbRelationDeleteCascade)
    Set NoMacam = Relasi1.CreateField("NoMacam")
    NoMacam.ForeignName = "NoMacam"
    Relasi1.Fields.Append NoMacam
    dbMesin.Relations.Append Relasi1
    
    'Membuat relasi antar table Jenis dan Relasi1
    Set Relasi2 = dbMesin.CreateRelation("Relasi2", "tblJenis", "tblRelasi1", dbRelationLeft + dbRelationDeleteCascade)
    Set NoJenis = Relasi2.CreateField("NoJenis")
    NoJenis.ForeignName = "NoJenis"
    Relasi2.Fields.Append NoJenis
    dbMesin.Relations.Append Relasi2
    
    'Membuat relasi table ciri dan relasi2
    Set Relasi3 = dbMesin.CreateRelation("Relasi3", "tblCiri", "tblRelasi2", dbRelationLeft + dbRelationDeleteCascade)
    Set NoCiri = Relasi3.CreateField("NoCiri")
    NoCiri.ForeignName = "NoCiri"
    Relasi3.Fields.Append NoCiri
    dbMesin.Relations.Append Relasi3
    
    'Membuat relasi Jenis dan relasi2
    Set Relasi4 = dbMesin.CreateRelation("Relasi4", "tblJenis", "tblRelasi2", dbRelationLeft + dbRelationDeleteCascade)
    Set NoJenis = Relasi4.CreateField("NoJenis")
    NoJenis.ForeignName = "NoJenis"
    Relasi4.Fields.Append NoJenis
    dbMesin.Relations.Append Relasi4
End Sub
