VERSION 5.00
Begin VB.Form frmSolusi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solusi Kerusakan Mesin"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelesai 
      Caption         =   "Selesai"
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
      Left            =   3353
      TabIndex        =   0
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtSolusi 
      Height          =   3495
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   7695
   End
End
Attribute VB_Name = "frmSolusi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSelesai_Click()
    Unload frmRekamanData
    Unload frmPenelusuranMacam
    Unload frmPenelusuranJenis
    Unload frmPenelusuranCiri
    Unload frmPenelusuran
    Unload Me
End Sub
