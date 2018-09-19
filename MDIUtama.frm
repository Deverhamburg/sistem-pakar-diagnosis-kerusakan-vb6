VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIUtama 
   Appearance      =   0  'Flat
   BackColor       =   &H80000011&
   Caption         =   "Menu Utama"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdlHelp 
      Left            =   2640
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2655
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "26/04/2017"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "12:00"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   1429
      ButtonWidth     =   2037
      ButtonHeight    =   1429
      Style           =   1
      ImageList       =   "imgIkon"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pengetahuan"
            Key             =   "Pengetahuan"
            Object.ToolTipText     =   "Menambahkan Basis Pengetahuan"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Macam"
                  Text            =   "Macam Kerusakan"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Jenis"
                  Text            =   "Jenis Kerusakan"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Ciri"
                  Text            =   "Ciri Kerusakan"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Penelusuran"
            Key             =   "Penelusuran"
            Object.ToolTipText     =   "Penelusuran Data"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Penjelasan"
            Key             =   "Penjelasan"
            Object.ToolTipText     =   "Penjelasan Sistem"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIkon 
      Left            =   2040
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIUtama.frx":0000
            Key             =   "Pengetahuan"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIUtama.frx":0452
            Key             =   "Penelusuran"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIUtama.frx":08A4
            Key             =   "Penjelasan"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPakar 
      Caption         =   "&Pakar"
      Begin VB.Menu mnuPengetahuan 
         Caption         =   "Basis Pen&getahuan"
         Begin VB.Menu mnuMacam 
            Caption         =   "Macam Kerusakan"
         End
         Begin VB.Menu mnuJenis 
            Caption         =   "Jenis Kerusakan"
         End
         Begin VB.Menu mnuCiri 
            Caption         =   "Ciri Kerusakan"
         End
      End
      Begin VB.Menu mnuAturan 
         Caption         =   "Basis A&turan"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuPenjelasan 
         Caption         =   "&Penjelasan Sistem"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuGaris 
         Caption         =   ""
      End
      Begin VB.Menu mnuKeluar 
         Caption         =   "K&eluar"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuPemakai 
      Caption         =   "Pe&makai"
      Begin VB.Menu mnuPenelusuran 
         Caption         =   "Penelu&suran"
      End
      Begin VB.Menu mnuPenjelasan2 
         Caption         =   "&Penjelasan Sistem"
      End
      Begin VB.Menu mnuGaris2 
         Caption         =   ""
      End
      Begin VB.Menu mnuSelesai 
         Caption         =   "S&elesai"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuManual 
         Caption         =   "Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuKeterangan 
         Caption         =   "Keterangan Program"
      End
   End
End
Attribute VB_Name = "MDIUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    Me.imgIkon.ListImages.Remove "Pengetahuan"
    Me.imgIkon.ListImages.Remove "Penelusuran"
    Me.imgIkon.ListImages.Remove "Penjelasan"
    Me.imgIkon.ListImages.Clear
    
    Me.imgIkon.ListImages.Add , "Pengetahuan", LoadPicture(App.Path + "\images\ikon\FOLDER05.ico")
    Me.imgIkon.ListImages.Add , "Penelusuran", LoadPicture(App.Path + "\images\ikon\CRDFLE13.ico")
    Me.imgIkon.ListImages.Add , "Penjelasan", LoadPicture(App.Path + "\images\ikon\CRDFLE03.ico")
    
    App.HelpFile = App.Path & "\help.chm"
End Sub

Private Sub mnuAturan_Click()
    Load frmAturan
    frmAturan.SetFocus
End Sub

Private Sub mnuManual_Click()
    Me.cdlHelp.HelpFile = App.Path & "\help.chm"
    Me.cdlHelp.HelpCommand = cdlHelpContext
    Me.cdlHelp.ShowHelp
End Sub

Private Sub mnuSelesai_Click()
    Call mnuKeluaran_click
End Sub

Private Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Penelusuran"
            Call mnuPenelusuran_click
        Case "Penjelasan"
            If Me.mnuPakar.Visible = True Then
                Call mnuPenjelasan_Click
            Else
                Call mnuPenjelasan2_Click
            End If
    End Select
End Sub

Private Sub toolbar_ButtonMenuCLick(ByVal buttonMenu As MSComctlLib.buttonMenu)
    Select Case buttonMenu.Key
        Case "Macam"
            mnuMacam_click
        Case "Jenis"
            mnuJenis_click
        Case "Ciri"
            mnuCiri_click
    End Select
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuPengetahuan
    End If
End Sub

Private Sub mnuMacam_click()
    Load frmMacamKerusakan
    frmMacamKerusakan.SetFocus
End Sub

Private Sub mnuCiri_click()
    Load frmCiriKerusakan
    frmCiriKerusakan.SetFocus
End Sub

Private Sub mnuJenis_click()
    Load frmJenisKerusakan
    frmJenisKerusakan.SetFocus
End Sub

Private Sub mnuKeterangan_click()
    Load frmKeterangan
    frmKeterangan.SetFocus
End Sub

Private Sub mnuKeluaran_click()
    Unload MDIUtama
End Sub

Private Sub mnuPenelusuran_click()
    Dim i As Integer
    Load frmPenelusuran
    For i = 0 To 2
        frmPenelusuran.chkPilihan(i).Value = vbChecked
    Next i
    frmPenelusuran.SetFocus
End Sub

Private Sub mnuPenjelasan_Click()
    Load frmPenjelasan
    frmPenjelasan.cmdSimpan.Visible = True
    frmPenjelasan.SetFocus
End Sub

Private Sub mnuPenjelasan2_Click()
    Load frmPenjelasan
    frmPenjelasan.cmdSimpan.Visible = False
    frmPenjelasan.SetFocus
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Unload frmCiriKerusakan
    Unload frmJenisKerusakan
    Unload frmKeterangan
    Unload frmMacamKerusakan
    
    Call mdlTutup.Tutup
End Sub
