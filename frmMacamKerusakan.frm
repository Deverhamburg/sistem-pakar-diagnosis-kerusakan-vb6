VERSION 5.00
Begin VB.Form frmMacamKerusakan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Macam Kerusakan Mesin"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTutup 
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
      Left            =   3960
      TabIndex        =   5
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "Hapus"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit Data"
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
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
   End
   Begin VB.ListBox lstMacam 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   5535
   End
   Begin VB.CommandButton cmdTambah 
      Appearance      =   0  'Flat
      Caption         =   "Tambahkan"
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
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtMacam 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label lblJudul 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Macam-Macam Kerusakan Mesin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3420
   End
End
Attribute VB_Name = "frmMacamKerusakan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEdit_click()
    Dim EditData As String
    Dim i As Integer
    Dim Kodenya As String
    Dim Datanya As String
    
    Kodenya = Mid(Me.lstMacam.Text, 1, 4)
    Datanya = Mid(Me.lstMacam.Text, 8, Len(Me.lstMacam.Text))
    
    If Datanya <> "" Then
        EditData = InputBox("Masukan data yang baru dari data", "Konfirmasi", Datanya)
        If EditData <> "" Then
            'Menentukan indexs
            tblMacam.index = "idMacam"
            tblMacam.Seek "=", Kodenya
            tblMacam.Edit
            tblMacam!Macam = EditData
            tblMacam.Update
            Call Form_Load
        End If
    End If
End Sub

Private Sub cmdHapus_click()
    Dim hapus As String
    Dim i As String
    'mengambil kode macam kerusakan
    hapus = Mid(Me.lstMacam.Text, 1, 4)
    If hapus <> "" Then
        If MsgBox("Apakah Anda yakin akan menghapus data " & hapus & " ?", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
            'Menentukan index
            tblMacam.index = "idMacam"
            'pencarian data
            tblMacam.Seek "=", hapus
            'menghapus data
            tblMacam.Delete
            Call Form_Load
        End If
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.lstMacam.Clear
    
    Dim i As Integer
    
    If tblMacam.RecordCount <> 0 Then
        tblMacam.MoveFirst
        For i = 1 To tblMacam.RecordCount
            Me.lstMacam.AddItem tblMacam!NoMacam & "   " & tblMacam!Macam
            tblMacam.MoveNext
        Next i
    End If
End Sub

Private Sub cmdTambah_click()
    Dim jawab As Integer
    Dim Ambil As String
    
    If Me.txtMacam.Text <> "" Then
        'ID generator
        Ambil = Me.lstMacam.List(lstMacam.ListCount - 1)
        Ambil = Mid(Ambil, 1, 4)
        Ambil = Right(Ambil, 3)
        Ambil = Val(Ambil + 1001)
        Ambil = Str(Ambil)
        Ambil = Right(Ambil, 3)
        Ambil = "M" & Ambil
        
        'Menyimpan data dalam table tblMacam
        tblMacam.MoveLast
        tblMacam.AddNew
        tblMacam!NoMacam = Ambil
        tblMacam!Macam = Me.txtMacam.Text
        tblMacam.Update
        
        Me.lstMacam.AddItem Ambil & "   " & Me.txtMacam.Text
        Me.txtMacam.Text = ""
        Me.txtMacam.SetFocus
    Else
        jawab = MsgBox("Anda belum menginputkan macam kerusakan !" & vbNewLine & "Silahkan masukan macam kerusakan terlebih dahulu !", vbOKOnly + vbCritical, "Konfirmasi")
        If jawab = vbOK Then
            Me.txtMacam.SetFocus
        End If
    End If
End Sub
