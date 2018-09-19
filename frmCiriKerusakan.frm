VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCiriKerusakan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ciri Kerusakan"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Begin VB.Data dtCiri 
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
      RecordSource    =   "tblCiri"
      Top             =   4440
      Visible         =   0   'False
      Width           =   2655
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
   Begin VB.TextBox txtCiri 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox txtDiagnosa 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   4935
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
   Begin MSFlexGridLib.MSFlexGrid grdCiri 
      Bindings        =   "frmCiriKerusakan.frx":0000
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
   Begin VB.Label lblNomerCiri 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nomer/Ciri Kerusakan"
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
      Top             =   600
      Width           =   2280
   End
   Begin VB.Label lblDiagnosa 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Diagnosa"
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
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label lblJudul 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ciri - Ciri Kerusakan Mesin"
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
      Width           =   2730
   End
End
Attribute VB_Name = "frmCiriKerusakan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAtas_click()
    tblCiri.MoveFirst
    Call Me.TampilData
End Sub

Private Sub cmdBawah_Click()
    tblCiri.MoveLast
    Call Me.TampilData
End Sub

Private Sub cmdSebelum_Click()
    tblCiri.MovePrevious
    If tblCiri.BOF Then
        tblCiri.MoveFirst
    End If
    Call Me.TampilData
End Sub

Private Sub cmdSesudah_click()
    tblCiri.MoveNext
    If tblCiri.EOF Then
        tblCiri.MoveLast
    End If
    Call Me.TampilData
End Sub

Sub TampilData()
    Me.txtNomer.Text = tblCiri!NoCiri
    Me.txtCiri.Text = tblCiri!Ciri
    Me.txtDiagnosa.Text = tblCiri!Diagnosa
End Sub

Private Sub Form_Load()
    Call cmdAtas_click
End Sub

Private Sub cmdTambah_click()
    Dim Ambil As String
    
    If tblCiri.RecordCount <> 0 Then
        tblCiri.MoveLast
        Ambil = tblCiri!NoCiri
        Ambil = Right(Ambil, 3)
        Ambil = Val(Ambil) + 1001
        Ambil = Str(Ambil)
        Ambil = Right(Ambil, 3)
        Ambil = "C" & Ambil
        Me.txtNomer.Text = Ambil
        Me.txtCiri.Text = ""
        Me.txtDiagnosa.Text = ""
        Me.txtCiri.Locked = False
        Me.txtDiagnosa.Locked = False
        Me.txtCiri.SetFocus
    End If
    tblCiri.AddNew
End Sub

Private Sub cmdEdit_click()
    Me.txtCiri.Locked = False
    Me.txtDiagnosa.Locked = False
    tblCiri.Edit
End Sub

Private Sub cmdSimpan_Click()
    tblCiri!NoCiri = Me.txtNomer.Text
    tblCiri!Ciri = Me.txtCiri.Text
    tblCiri!Diagnosa = Me.txtDiagnosa.Text
    tblCiri.Update
    Me.txtCiri.Locked = True
    Me.txtDiagnosa.Locked = True
    Call Me.tampil_grid
End Sub

Private Sub cmdHapus_click()
    If MsgBox("Apakah anda yakin untuk menghapus data" & Me.txtCiri.Text & " ?", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
        tblCiri.Delete
        Call cmdSebelum_Click
        Call Me.tampil_grid
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Sub tampil_grid()
    Dim i As Integer
    
    Me.dtCiri.Refresh
    Me.grdCiri.Refresh
    
    Me.grdCiri.ColWidth(0) = 800
    Me.grdCiri.ColWidth(1) = 4000
    Me.grdCiri.ColWidth(2) = 0
    
    Me.grdCiri.Row = 0
    For i = 0 To Me.grdCiri.Col - 1
        Me.grdCiri.Col = i
        Me.grdCiri.CellFontBold = True
        Me.grdCiri.CellAlignment = flexAlignCenterCenter
    Next i
End Sub

Private Sub form_activate()
    Call Me.tampil_grid
End Sub
