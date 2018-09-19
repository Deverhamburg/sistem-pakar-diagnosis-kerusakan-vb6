VERSION 5.00
Begin VB.Form frmPenelusuranJenis 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Penelusuran Jenis Kerusakan Mesin"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdKembali 
      Caption         =   "Kembali"
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
      Left            =   694
      TabIndex        =   4
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox txtGejala 
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmPenelusuranJenis.frx":0000
      Top             =   3360
      Width           =   6015
   End
   Begin VB.ListBox lstJenis 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   6015
   End
   Begin VB.TextBox txtMacam 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   6015
   End
   Begin VB.CommandButton cmdLanjutkan 
      Caption         =   "Lanjutkan"
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
      Left            =   3581
      TabIndex        =   3
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label lbl3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Gejala jenis kerusakan jenis di atas adalah :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   6015
   End
   Begin VB.Label lbl2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih jenis kerusakan dari macam kerusakan tersebut di atas :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label lbl1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Anda memilih macam kerusakan mesin sebagai berikut :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "frmPenelusuranJenis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdKembali_Click()
    Unload Me
    frmPenelusuranMacam.SetFocus
End Sub

Private Sub cmdLanjutkan_Click()
    Load frmPenelusuranCiri
    frmPenelusuranCiri.txtJenis.Text = Me.lstJenis.Text
    frmPenelusuranCiri.SetFocus
End Sub

Private Sub form_activate()
    Dim Ambil As String
    Dim i As Integer
    
    Me.lstJenis.Clear
    Ambil = Mid(Me.txtMacam.Text, 1, 4)
    If tblRelasi1.RecordCount <> 0 Then
        tblRelasi1.MoveFirst
        Do While Not tblRelasi1.EOF
            If tblRelasi1!NoMacam = Ambil Then
                tblJenis.index = "idJenis"
                tblJenis.Seek "=", tblRelasi1!NoJenis
                Me.lstJenis.AddItem tblJenis!NoJenis & "   " & tblJenis!Jenis
            End If
            tblRelasi1.MoveNext
        Loop
        If Me.lstJenis.ListCount <> 0 Then
            Me.lstJenis.ListIndex = 0
            Call lstJenis_click
        End If
    End If
    
    MDIUtama.cdlHelp.HelpContext = 2
End Sub

Private Sub lstJenis_click()
    Dim Ambil As String
    Ambil = Mid(Me.lstJenis.Text, 1, 4)
    tblJenis.index = "idJenis"
    tblJenis.Seek "=", Ambil
    Me.txtGejala.Text = tblJenis!Gejala
End Sub

