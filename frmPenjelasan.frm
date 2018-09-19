VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPenjelasan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Penjelasan Sistem"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSimpan 
      Appearance      =   0  'Flat
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
      Left            =   5160
      TabIndex        =   2
      Top             =   4920
      Width           =   1455
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
      Left            =   6840
      TabIndex        =   1
      Top             =   4920
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox rtfPenjelasan 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7646
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmPenjelasan.frx":0000
   End
End
Attribute VB_Name = "frmPenjelasan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub form_activate()
    Me.rtfPenjelasan.LoadFile App.Path & "\penjelasan.rtf"
End Sub

Private Sub cmdSimpan_Click()
    Me.rtfPenjelasan.SaveFile App.Path & "\penjelasan.rtf"
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub
