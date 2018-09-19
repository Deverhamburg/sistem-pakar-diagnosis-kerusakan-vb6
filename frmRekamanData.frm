VERSION 5.00
Begin VB.Form frmRekamanData 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rekaman data yang akan dianalisis"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8190
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
      Left            =   752
      TabIndex        =   2
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox txtRekaman 
      Height          =   3495
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   7695
   End
   Begin VB.CommandButton cmdProses 
      Caption         =   "Proses"
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
      Left            =   4904
      TabIndex        =   0
      Top             =   3840
      Width           =   2535
   End
End
Attribute VB_Name = "frmRekamanData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdKembali_Click()
    Unload Me
    frmPenelusuranCiri.SetFocus
End Sub

Private Sub cmdProses_Click()
    Dim Ambil As String
    Ambil = Mid(frmPenelusuranCiri.lstCiri.Text, 1, 4)
    tblCiri.index = "idCiri"
    tblCiri.Seek "=", Ambil
    Load frmAnimasi
End Sub

Private Sub form_activate()
    Dim Keterangan As String
    
    Keterangan = "Sistem sudah merekam data yang anda pilih yaitu :"
    Keterangan = Keterangan & vbNewLine
    Keterangan = Keterangan & "Macam Kerusakan Mesin, Jenis dan Cirinya"
    Keterangan = Keterangan & vbNewLine
    Keterangan = Keterangan & "Data yang terekam berturut-turut adalah sebagai berikut :"
    Keterangan = Keterangan & vbNewLine
    Keterangan = Keterangan & frmPenelusuranMacam.cmbPenelusuranMacam.Text
    Keterangan = Keterangan & vbNewLine
    Keterangan = Keterangan & frmPenelusuranJenis.lstJenis.Text
    Keterangan = Keterangan & vbNewLine
    Keterangan = Keterangan & frmPenelusuranCiri.lstCiri.Text
    Me.txtRekaman.Text = Keterangan
End Sub

