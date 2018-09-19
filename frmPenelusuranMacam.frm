VERSION 5.00
Begin VB.Form frmPenelusuranMacam 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Penelusuran Macam Kerusakan Mesin"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4485
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
      Left            =   195
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.ComboBox cmbPenelusuranMacam 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   4215
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
      Left            =   2475
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblPertanyaan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kerusakan mesin apa yang anda temukan ?"
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
      TabIndex        =   3
      Top             =   720
      Width           =   4215
   End
End
Attribute VB_Name = "frmPenelusuranMacam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdKembali_Click()
    Unload Me
    frmPenelusuran.SetFocus
End Sub

Public Sub form_activate()
    Dim i As Integer
    Me.cmbPenelusuranMacam.Clear
    If tblMacam.RecordCount <> 0 Then
        tblMacam.MoveFirst
        Do While Not tblMacam.EOF
            Me.cmbPenelusuranMacam.AddItem tblMacam!NoMacam & "   " & tblMacam!Macam
            tblMacam.MoveNext
        Loop
        Me.cmbPenelusuranMacam.ListIndex = 0
    End If
    
    MDIUtama.cdlHelp.HelpContext = 1
End Sub

Private Sub cmdLanjutkan_Click()
    Load frmPenelusuranJenis
    frmPenelusuranJenis.txtMacam.Text = Me.cmbPenelusuranMacam.Text
    frmPenelusuranJenis.SetFocus
End Sub

Private Sub Form_Load()
    Me.cmbPenelusuranMacam.Clear
    With Me.cmbPenelusuranMacam
        .AddItem "Mesin"
        .AddItem "Transmisi Daya"
        .AddItem "Sistem Kemudi"
        .AddItem "Sistem Suspensi"
        .AddItem "Roda"
        .AddItem "Rem"
        .AddItem "Lampu"
        .AddItem "Klakson"
    End With
End Sub
