VERSION 5.00
Begin VB.Form frmTambahAturan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menambahkan Aturan"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbMacam 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.ListBox lstCiri 
      Appearance      =   0  'Flat
      Height          =   1605
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   3960
      Width           =   5775
   End
   Begin VB.ListBox lstJenis 
      Appearance      =   0  'Flat
      Height          =   1605
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1320
      Width           =   5775
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
      Left            =   4920
      TabIndex        =   5
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSimpanJenis 
      Caption         =   "Simpan Jenis dan Ciri Kerusakan"
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
      Left            =   240
      TabIndex        =   4
      Top             =   5640
      Width           =   5775
   End
   Begin VB.CommandButton cmdSimpanMacam 
      Appearance      =   0  'Flat
      Caption         =   "Simpan Macam dan Jenis Kerusakan"
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
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   5775
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
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   1860
   End
   Begin VB.Label lblJenis 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis-Jenis Kerusakan"
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
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   2355
   End
   Begin VB.Label lblMacam 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Macam Kerusakan"
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
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1920
   End
End
Attribute VB_Name = "frmTambahAturan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub TampilJenis()
    Dim Ambil As String
    Dim i As Integer
    Dim j As Integer
    
    For j = 0 To Me.lstJenis.ListCount - 1
        Me.lstJenis.Selected(j) = False
    Next j
    
    Ambil = Mid(Me.cmbMacam.Text, 1, 4)
    tblRelasi1.MoveFirst
    Do While Not tblRelasi1.EOF
    'For i = 1 To tblRelasi1.RecordCount
        If tblRelasi1!NoMacam = Ambil Then
            For j = 0 To Me.lstJenis.ListCount - 1
                If Mid(Me.lstJenis.List(j), 1, 4) = tblRelasi1!NoJenis Then
                    Me.lstJenis.Selected(j) = True
                End If
            Next j
        End If
        tblRelasi1.MoveNext
    'Next i
    Loop
End Sub

Sub TampilCiri()
    Dim Ambil As String
    Dim i As Integer
    Dim j As Integer
    
    For j = 0 To Me.lstCiri.ListCount - 1
        Me.lstCiri.Selected(j) = False
    Next j
    
    Ambil = Mid(Me.lstJenis.Text, 1, 4)
    tblRelasi2.MoveFirst
    Do While Not tblRelasi2.EOF
    'For i = 1 To tblRelasi2.RecordCount
        If tblRelasi2!NoJenis = Ambil Then
            For j = 0 To Me.lstCiri.ListCount - 1
                If Mid(Me.lstCiri.List(j), 1, 4) = tblRelasi2!NoCiri Then
                    Me.lstCiri.Selected(j) = True
                End If
            Next j
        End If
        tblRelasi2.MoveNext
    'Next i
    Loop
End Sub

Private Sub form_activate()
    Dim i As Integer
    Me.cmbMacam.Clear
    If tblMacam.RecordCount <> 0 Then
        tblMacam.MoveFirst
        For i = 1 To tblMacam.RecordCount
            Me.cmbMacam.AddItem tblMacam!NoMacam & "   " & tblMacam!Macam
            tblMacam.MoveNext
        Next i
        Me.cmbMacam.ListIndex = 0
    End If
    
    Me.lstJenis.Clear
    If tblJenis.RecordCount <> 0 Then
        tblJenis.MoveFirst
        For i = 1 To tblJenis.RecordCount
            Me.lstJenis.AddItem tblJenis!NoJenis & "   " & tblJenis!Jenis
            tblJenis.MoveNext
        Next i
        Me.lstJenis.ListIndex = 0
    End If
    
    Me.lstCiri.Clear
    If tblCiri.RecordCount <> 0 Then
        tblCiri.MoveFirst
        For i = 1 To tblCiri.RecordCount
            Me.lstCiri.AddItem tblCiri!NoCiri & "   " & tblCiri!Ciri
        Next i
    End If
    Me.lstCiri.ListIndex = 0
End Sub

Private Sub cmdSimpanJenis_click()
    Dim Ambil As String
    Dim i As Integer
    
    Ambil = Mid(Me.lstJenis.Text, 1, 4)
    If tblRelasi2.RecordCount <> 0 Then
        tblRelasi2.MoveFirst
        Do While Not tblRelasi2.EOF
        'For i = 1 To tblRelasi2.RecordCount
            If tblRelasi2!NoJenis = Ambil Then
                tblRelasi2.Delete
            End If
            tblRelasi2.MoveNext
        'Next i
        Loop
    End If
    
    For i = 0 To Me.lstCiri.ListCount - 1
        If Me.lstCiri.Selected(i) = True Then
            tblRelasi2.AddNew
            tblRelasi2!NoJenis = Ambil
            tblRelasi2!NoCiri = Mid(Me.lstCiri.List(i), 1, 4)
            tblRelasi2.Update
        End If
    Next i
End Sub

Private Sub cmdSimpanMacam_click()
    Dim Ambil As String
    Dim i As Integer
    
    Ambil = Mid(Me.cmbMacam.Text, 1, 4)
    If tblRelasi1.RecordCount <> 0 Then
        tblRelasi1.MoveFirst
        Do While Not tblRelasi1.EOF
        'For i = 1 To tblRelasi1.RecordCount
            If tblRelasi1!NoMacam = Ambil Then
                tblRelasi1.Delete
            End If
            tblRelasi1.MoveNext
        'Next i
        Loop
    End If
    
    For i = 0 To Me.lstJenis.ListCount - 1
        If Me.lstJenis.Selected(i) = True Then
            tblRelasi1.AddNew
            tblRelasi1!NoMacam = Ambil
            tblRelasi1!NoJenis = Mid(Me.lstJenis.List(i), 1, 4)
            tblRelasi1.Update
        End If
    Next i
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub lstJenis_click()
    Call Me.TampilCiri
End Sub

Private Sub cmbMacam_click()
    Call Me.TampilJenis
End Sub
