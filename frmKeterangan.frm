VERSION 5.00
Begin VB.Form frmKeterangan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Keterangan"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   240
      Top             =   2520
   End
   Begin VB.CommandButton cmdTutup 
      Appearance      =   0  'Flat
      Caption         =   "Tutup"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblSelamatDatang 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Selamat Datang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1470
      TabIndex        =   1
      Top             =   600
      Width           =   1965
   End
   Begin VB.Label lblSistemPakar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Di Program Sistem Pakar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   945
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FF80FF&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "frmKeterangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTutup_click()
Unload Me
Set frmKeterangan = Nothing
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Width = 5000
Me.Height = 3600
End Sub

Private Sub Timer1_Timer()
    Me.lblSelamatDatang.Caption = Mid(Me.lblSelamatDatang.Caption, 2, Len(Me.lblSelamatDatang.Caption)) + Mid(Me.lblSelamatDatang.Caption, 1, 1)
    If Me.lblSistemPakar.ForeColor = vbWindowText Then
        Me.lblSistemPakar.ForeColor = vbHighlightText
    Else
        Me.lblSistemPakar.ForeColor = vbWindowText
    End If
End Sub
