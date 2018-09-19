VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAnimasi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proses..."
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   3840
      Top             =   720
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Max             =   10000
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1508
      _Version        =   393216
      FullWidth       =   273
      FullHeight      =   57
   End
End
Attribute VB_Name = "frmAnimasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.ProgressBar1.Max = 10000
    Me.ProgressBar1.Min = 0
    Me.Timer1.Interval = 150
    
    Me.Animation1.Open App.Path + "\videos\FILECOPY.AVI"
    Me.Animation1.Play
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    For i = Me.ProgressBar1.Min To Me.ProgressBar1.Max
        Me.ProgressBar1.Value = i
        If Me.ProgressBar1.Value = 10000 Then
            Unload Me
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Animation1.Stop
    Me.Timer1.Interval = 0
    Load frmSolusi
    
    If Not tblCiri.EOF Then
        frmSolusi.txtSolusi.Text = tblCiri!Diagnosa
    Else
        frmSolusi.txtSolusi.Text = ""
    End If
    
    frmSolusi.SetFocus
End Sub

