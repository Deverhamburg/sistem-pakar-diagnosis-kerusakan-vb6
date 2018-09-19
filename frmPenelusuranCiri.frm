VERSION 5.00
Begin VB.Form frmPenelusuranCiri 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Penulusuran Ciri Kerusakan Mesin"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6165
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
      Left            =   667
      TabIndex        =   2
      Top             =   4440
      Width           =   2055
   End
   Begin VB.ListBox lstCiri 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   5655
   End
   Begin VB.TextBox txtJenis 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   5655
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
      Left            =   3442
      TabIndex        =   3
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label lbl2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih ciri kerusakan dari jenis kerusakan tersebut di atas :"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Label lbl1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Anda memilih jenis kerusakan mesin sebagai sebagai berikut :"
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
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   5265
   End
End
Attribute VB_Name = "frmPenelusuranCiri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdKembali_Click()
    Unload Me
    frmPenelusuranJenis.SetFocus
End Sub

Private Sub cmdLanjutkan_Click()
    Load frmRekamanData
    frmRekamanData.SetFocus
End Sub

Private Sub form_activate()
    Dim Ambil As String
    Dim i As Integer
    
    Me.lstCiri.Clear
    Ambil = Mid(Me.txtJenis.Text, 1, 4)
    If tblRelasi2.RecordCount <> 0 Then
        tblRelasi2.MoveFirst
        Do While Not tblRelasi2.EOF
            If tblRelasi2!NoJenis = Ambil Then
                tblCiri.index = "idCiri"
                tblCiri.Seek "=", tblRelasi2!NoCiri
                Me.lstCiri.AddItem tblCiri!NoCiri & "   " & tblCiri!Ciri
            End If
            tblRelasi2.MoveNext
        Loop
        If Me.lstCiri.ListCount <> 0 Then
            Me.lstCiri.ListIndex = 0
        End If
    End If
    
    MDIUtama.cdlHelp.HelpContext = 3
End Sub
