VERSION 5.00
Begin VB.Form frmAturan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aturan"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTambah 
      Appearance      =   0  'Flat
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdTutup 
      Appearance      =   0  'Flat
      Caption         =   "Tutup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox txtDiagnosa 
      Appearance      =   0  'Flat
      Height          =   2000
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3960
      Width           =   4275
   End
   Begin VB.TextBox txtGejala 
      Appearance      =   0  'Flat
      Height          =   2000
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   4275
   End
   Begin VB.ListBox lstCiri 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   3255
   End
   Begin VB.ListBox lstJenis 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   3255
   End
   Begin VB.ComboBox cmbMacam 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblDiagnosa 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Diagnosa Kerusakaan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3720
      TabIndex        =   11
      Top             =   3600
      Width           =   2310
   End
   Begin VB.Label lblCiri 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ciri-ciri Kerusakaan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   2040
   End
   Begin VB.Label lblGejala 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Gejala Kerusakan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3720
      TabIndex        =   9
      Top             =   1080
      Width           =   1860
   End
   Begin VB.Label lblJenisKerusakan 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis-Jenis Kerusakan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   2355
   End
   Begin VB.Label lblMacamKerusakan 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Macam Kerusakan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1920
   End
End
Attribute VB_Name = "frmAturan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTambah_click()
    Load frmTambahAturan
    frmTambahAturan.SetFocus
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Sub TampilkanJenis()
    Dim Ambil As String
    Dim cari As String
    Dim i As Integer
    
    Me.lstJenis.Clear
    Me.txtGejala.Text = ""
    Ambil = Mid(Me.cmbMacam.Text, 1, 4)
    If tblRelasi1.RecordCount <> 0 Then
        tblRelasi1.MoveFirst
        Do While Not tblRelasi1.EOF
        'For i = 1 To tblRelasi1.RecordCount
            If tblRelasi1!NoMacam = Ambil Then
                cari = tblRelasi1!NoJenis
                tblJenis.index = "idJenis"
                tblJenis.Seek "=", cari
                Me.lstJenis.AddItem tblJenis!NoJenis & "   " & tblJenis!Jenis
            End If
            tblRelasi1.MoveNext
            Call Me.TampilkanCiri
            If Me.lstJenis.ListCount <> 0 Then
                Me.lstJenis.ListIndex = 0
                Call lstJenis_click
            End If
        'Next i
        Loop
    End If
End Sub

Sub TampilkanCiri()
    Dim Ambil As String
    Dim cari As String
    Dim i As Integer
    
    Me.txtDiagnosa.Text = ""
    Me.lstCiri.Clear
    Ambil = Mid(Me.lstJenis.Text, 1, 4)
    If tblRelasi2.RecordCount <> 0 Then
        tblRelasi2.MoveFirst
        Do While Not tblRelasi2.EOF
        'For i = 1 To tblRelasi2.RecordCount
           If tblRelasi2!NoJenis = Ambil Then
                cari = tblRelasi2!NoCiri
                tblCiri.index = "idCiri"
                tblCiri.Seek "=", cari
                Me.lstCiri.AddItem tblCiri!NoCiri & "   " & tblCiri!Ciri
            End If
            tblRelasi2.MoveNext
            If Me.lstCiri.ListCount <> 0 Then
                Me.lstCiri.ListIndex = 0
                Call lstCiri_Click
            End If
        'next i
        Loop
    End If
End Sub

Private Sub form_activate()
    Dim i As Integer
    
    Me.cmbMacam.Clear
    If tblMacam.RecordCount <> 0 Then
        tblMacam.MoveFirst
        For i = 1 To tblMacam.RecordCount
            Me.cmbMacam.AddItem tblMacam!NoMacam & "   " & tblMacam!Macam
            tblMacam.MoveNext
        Next i
        Me.cmbMacam.ListIndex = 0
        Call Me.TampilkanJenis
    End If
End Sub

Private Sub cmbMacam_click()
    Call Me.TampilkanJenis
End Sub

Private Sub lstCiri_Click()
    Dim cari As String
    Dim i As Integer
    
    cari = Mid(Me.lstCiri.Text, 1, 4)
    tblCiri.index = "idCiri"
    tblCiri.Seek "=", cari
    Me.txtDiagnosa.Text = tblCiri!Diagnosa
End Sub

Private Sub lstJenis_click()
    Dim cari As String
    Dim i As Integer
    
    cari = Mid(Me.lstJenis.Text, 1, 4)
    tblJenis.index = "idJenis"
    tblJenis.Seek "=", cari
    Me.txtGejala.Text = tblJenis!Gejala
    Call Me.TampilkanCiri
End Sub
