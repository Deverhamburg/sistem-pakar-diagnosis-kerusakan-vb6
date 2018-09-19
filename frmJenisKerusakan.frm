VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmJenisKerusakan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jenis Kerusakan"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8115
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
      Left            =   5880
      TabIndex        =   12
      Top             =   5160
      Width           =   1335
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
      Left            =   4680
      TabIndex        =   11
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simpan"
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
      Left            =   3480
      TabIndex        =   10
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
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
      Left            =   2280
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdTambah 
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
      Left            =   1080
      TabIndex        =   8
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdBawah 
      Caption         =   "> |"
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
      Left            =   1920
      TabIndex        =   7
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdSesudah 
      Caption         =   ">"
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
      Left            =   1320
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdSebelum 
      Caption         =   "<"
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
      Left            =   720
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdAtas 
      Caption         =   "| <"
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
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtGejala 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   4935
   End
   Begin VB.TextBox txtJenis 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox txtNomer 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Data dtJenis 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Projects\Sistem Pakar\program\dbMesin.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "tblJenis"
      Top             =   4440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid grdJenis 
      Bindings        =   "frmJenisKerusakan.frx":0000
      Height          =   1935
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   3413
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483643
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
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   1860
   End
   Begin VB.Label lblNomerJenis 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nomer/Jenis Kerusakan"
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
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   2490
   End
   Begin VB.Label lblJudul 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis-Jenis Kerusakan Mesin"
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
      TabIndex        =   13
      Top             =   120
      Width           =   3030
   End
End
Attribute VB_Name = "frmJenisKerusakan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub form_activate()
 Call Me.tampil_grid
End Sub

Sub tampil_grid()
    Dim i As Integer
    
    Me.dtJenis.Refresh
    Me.grdJenis.Refresh
    
    Me.grdJenis.ColWidth(0) = 800
    Me.grdJenis.ColWidth(1) = 4000
    Me.grdJenis.ColWidth(2) = 0
    
    Me.grdJenis.Row = 0
    For i = 0 To Me.grdJenis.Cols - 1
        Me.grdJenis.Col = i
        Me.grdJenis.CellFontBold = True
        Me.grdJenis.CellAlignment = flexAlignCenterCenter
    Next i
End Sub

Sub TampilData()
    Me.txtNomer.Text = tblJenis!NoJenis
    Me.txtJenis.Text = tblJenis!Jenis
    Me.txtGejala.Text = tblJenis!Gejala
End Sub

Private Sub cmdAtas_click()
    tblJenis.MoveFirst
    Call Me.TampilData
End Sub

Private Sub cmdBawah_Click()
    tblJenis.MoveLast
    Call Me.TampilData
End Sub

Private Sub cmdSebelum_Click()
    tblJenis.MovePrevious
    If tblJenis.BOF Then
        tblJenis.MoveFirst
    End If
    Call Me.TampilData
End Sub

Private Sub cmdSesudah_click()
    tblJenis.MoveNext
    If tblJenis.EOF Then
        tblJenis.MoveLast
    End If
    Call Me.TampilData
End Sub

Private Sub Form_Load()
    Me.txtJenis.Locked = True
    Me.txtGejala.Locked = True
    Call cmdAtas_click
End Sub

Private Sub cmdTambah_click()
    Dim Ambil As String
    If tblJenis.RecordCount <> 0 Then
        tblJenis.MoveLast
        Ambil = tblJenis!NoJenis
        Ambil = Right(Ambil, 3)
        Ambil = Val(Ambil) + 1001
        Ambil = Str(Ambil)
        Ambil = Right(Ambil, 3)
        Ambil = "J" & Ambil
        Me.txtNomer.Text = Ambil
        Me.txtJenis.Text = ""
        Me.txtGejala.Text = ""
        Me.txtJenis.Locked = False
        Me.txtGejala.Locked = False
        Me.txtJenis.SetFocus
    End If
    tblJenis.AddNew
End Sub

Private Sub cmdEdit_click()
    Me.txtJenis.Locked = False
    Me.txtGejala.Locked = False
    tblJenis.Edit
End Sub

Private Sub cmdSimpan_Click()
    tblJenis!NoJenis = Me.txtNomer.Text
    tblJenis!Jenis = Me.txtJenis.Text
    tblJenis!Gejala = Me.txtGejala.Text
    tblJenis.Update
    Me.txtJenis.Locked = True
    Me.txtGejala.Locked = True
    Call Me.tampil_grid
End Sub

Private Sub cmdHapus_click()
    If MsgBox("Apakah anda yakin akan menghapus data " & Me.txtJenis.Text & " ?", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
        tblJenis.Delete
        Call cmdSebelum_Click
        Call Me.tampil_grid
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub
