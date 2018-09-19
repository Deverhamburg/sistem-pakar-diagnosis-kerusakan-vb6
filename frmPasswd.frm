VERSION 5.00
Begin VB.Form frmPasswd 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPilihan 
      BackColor       =   &H80000005&
      Caption         =   "Pilihan Anda"
      Height          =   2775
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   4095
      Begin VB.CommandButton cmdGanti 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ganti"
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2040
         Width           =   1095
      End
      Begin VB.OptionButton optPilihan 
         BackColor       =   &H80000005&
         Caption         =   "Pakar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optPilihan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Pemakai"
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
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtPasswd 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtNama 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton cmdLogin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Login"
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
         Left            =   2700
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdTutup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblNama 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblPasswd 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Top             =   1320
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmPasswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGanti_Click()
    If Me.cmdGanti.Caption = "Ganti" Then
        If (Me.txtNama.Text = tblPasswd!Nama) And (Me.txtPasswd.Text = tblPasswd!Passwd) Then
            MsgBox "Silahkan memasukan nama dan password baru !", vbOKOnly + vbInformation, "Konfirmasi"
            Me.txtNama.Text = ""
            Me.txtPasswd.Text = ""
            Me.txtNama.SetFocus
            Me.cmdGanti.Caption = "Simpan"
        Else
            MsgBox "Masukan terlebih dahulu nama dan password anda yang lama dengan benar", vbOKOnly + vbInformation, "Konfirmasi"
            Exit Sub
        End If
    Else
        If (Me.txtNama.Text <> "") And (Me.txtPasswd.Text <> "") Then
            tblPasswd.MoveFirst
            tblPasswd.Edit
            tblPasswd!Nama = Me.txtNama.Text
            tblPasswd!Passwd = Me.txtPasswd.Text
            tblPasswd.Update
            MsgBox "Nama dan password anda yang baru siap digunakan !", vbOKOnly + vbInformation, "Konfirmasi"
            Me.txtNama.Text = ""
            Me.txtPasswd.Text = ""
            Me.txtNama.SetFocus
            Me.cmdGanti.Caption = "Ganti"
        Else
            MsgBox "Masukan terlebih dahulu nama dan password anda dengan benar", vbOKOnly + vbInformation, "Konfirmasi"
        End If
    End If
End Sub

Private Sub Form_Load()
    'cek jika .mdb ada
    If Dir(App.Path + "\dbMesin.mdb") <> "" Then
        mdlBuka.Buka
    Else
        If MsgBox("Kami kehilangan koneksi ke database, kemungkinan telah terjadi sesuatu pada program anda" & vbNewLine & "Program otomatis akan segera membuat database baru dan kemungkinan data anda sebelumnya akan hilang, apakah anda ingin melanjutkanya ?", vbYesNo + vbCritical, "Kehilangan koneksi dengan database") = vbYes Then
            mdlDB.buatDB
            mdlBuka.Buka
            buatDefaultUser
        Else
            MsgBox "Kami mengerti, maaf atas gangguan ini" & vbNewLine & "Silahkan hubungi administrasi anda untuk solusi terbaik.", vbOKOnly + vbInformation, "Kehilangan koneksi dengan database"
            Unload Me
        End If
    End If
End Sub

Sub buatDefaultUser()
    tblPasswd.AddNew
    tblPasswd!Nama = "admin"
    tblPasswd!Passwd = "123"
    tblPasswd.Update
    
    MsgBox "Untuk login dengan keamanan dan fasilitas Pakar silahkan anda login dengan :" & vbNewLine & "Nama      : admin" & vbNewLine & "Password  : 123", vbOKOnly + vbInformation, "Informasi untuk Login Pakar"
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    strValid = "abcdefghijklmnopqrstuvwxyz"
    strValid = strValid & "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    If KeyAscii = vbKeyReturn Then
        Me.txtPasswd.SetFocus
    Else
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtPasswd_keyPress(KeyAscii As Integer)
    Dim strValid As String
    strValid = "0123456789"
    If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNama_change()
    Me.txtNama.SelStart = Len(Me.txtNama.Text)
End Sub

Private Sub txtPasswd_change()
    If Len(Me.txtPasswd.Text) = 8 Then
        Me.cmdLogin.SetFocus
    End If
End Sub

'Private Sub txtNama_lostFocus()
'    If Me.txtNama.Text = "" Then
'        MsgBox "Anda harus menginputkan nama terlebih dahulu !", vbOKOnly + vbInformation, "Konfirmasi"
'        Me.txtNama.SetFocus
'    End If
'End Sub

Private Sub txtNama_gotfocus()
    Me.optPilihan(1).Value = True
End Sub

Private Sub optPilihan_click(index As Integer)
    If index = 1 Then
        Me.txtNama.SetFocus
    End If
End Sub

Private Sub cmdLogin_click()
    If Me.optPilihan(0).Value = True Then
        MDIUtama.mnuPakar.Visible = False
        MDIUtama.ToolBar.Buttons(1).Visible = False
        MDIUtama.StatusBar.Panels(3).Text = "Pemakai"
    Else
        tblPasswd.MoveFirst
        If (Me.txtNama.Text = tblPasswd!Nama) And (Me.txtPasswd.Text = tblPasswd!Passwd) Then
            MDIUtama.mnuPemakai.Visible = False
            MDIUtama.ToolBar.Buttons(2).Visible = False
            MDIUtama.StatusBar.Panels(3).Text = "Pakar"
        Else
            MsgBox "Password anda salah !", vbOKOnly + vbCritical, "Konfirmasi"
                Me.txtNama.Text = ""
                Me.txtPasswd.Text = ""
            Me.txtNama.SetFocus
            Exit Sub
        End If
    End If
    MDIUtama.Show
    MDIUtama.SetFocus
    Unload Me
End Sub

Private Sub cmdTutup_Click()
    End
End Sub
