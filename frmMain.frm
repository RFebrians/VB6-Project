VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Software Inventory"
   ClientHeight    =   2895
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   1   'True
   ScaleHeight     =   2895
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   0
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraControl 
      Height          =   975
      Left            =   600
      TabIndex        =   12
      Top             =   1800
      Width           =   4000
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtGo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>|"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">>"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<<"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<<"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraRecord 
      Caption         =   "Record ##"
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtCompany 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox txtSerial 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox txtTitle
         Alignment       =   2  'Center
         Caption         =   "Judul"
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFOpen 
         Caption         =   "&Open Database"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFSpace0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTNew 
         Caption         =   "&New Record"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuTEdit 
         Caption         =   "&Edit Record"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuTDelete 
         Caption         =   "&Delete Record"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuspace0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTPrint 
         Caption         =   "&View Print Records"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHHelp 
         Caption         =   "&Help"
         Enabled         =   0   'False
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHSpace0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------
' Deklarasi semua variable menggunakan 
' statement Option Explicit
' Bekerja mirip seperti import .
'--------------------------------------------------------------
Option Explicit
    Dim db As Database
    Dim rs As Recordset
    Dim DoWhat As Integer
    Dim NumRows As Integer

Private Sub Form_Load()
    '--------------------------------------------------------------
    ' Setup Database dan Object , 
    ' serta penggunaan dan memuat database itu sendiri
    '--------------------------------------------------------------
    Set db = OpenDatabase(App.Path & "\data.mdb")
    Set rs = db.OpenRecordset("tblSoftware", dbOpenDynaset)
    cmdSave.Enabled = False
    cmdClear.Enabled = False
    cmdCancel.Enabled = False
    cmdCancel.Visible = False
    txtTitle.Text = "Title"
    txtTitle.Locked = True
    txtSerial.Text = ""
    txtSerial.Locked = True
    txtCompany.Text = ""
    txtCompany.Locked = True
    fraRecord.Caption = "Record:"
    rs.MoveFirst
    GetData
End Sub

Private Sub cmdSave_Click()
    '--------------------------------------------------------------
    ' Interaksi CRUD (Create Read Update Delete)
    ' DoWhat akan menyatakan sebuah kondisi jika terdapat perubahan
    '--------------------------------------------------------------
    cmdSave.Enabled = False
    cmdClear.Enabled = False
    cmdGo.Enabled = True
    cmdFirst.Enabled = True
    cmdNext.Enabled = True
    cmdBack.Enabled = True
    cmdLast.Enabled = True
    txtGo.Enabled = True
    If DoWhat = 0 Then
        With rs
            .AddNew
                !Title = txtTitle.Text
                !Serial = txtSerial.Text
                !Company = txtCompany.Text
            .Update
        End With
    Else
        With rs
            .Edit
                !Title = txtTitle.Text
                !Serial = txtSerial.Text
                !Company = txtCompany.Text
            .Update
        End With
    End If
End Sub

Private Sub cmdClear_Click()
    '--------------------------------------------------------------
    ' Digunakan untuk menghapus teks
    '--------------------------------------------------------------
    cmdGo.Enabled = False
    cmdFirst.Enabled = False
    cmdNext.Enabled = False
    cmdBack.Enabled = False
    cmdLast.Enabled = False
    txtGo.Enabled = False
    If DoWhat = 0 Then
        txtTitle.Text = ""
        txtSerial.Text = ""
        txtCompany.Text = ""
    Else
        GetData
    End If
End Sub
    
Private Sub cmdCancel_Click()
    '--------------------------------------------------------------
    ' Mode batal saat Editing mode berlangsung 
    ' ini akan memuat ke View Mode
    '--------------------------------------------------------------
    cmdGo.Visible = True
    cmdGo.Enabled = True
    cmdFirst.Enabled = True
    cmdBack.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    Form_Load
End Sub

Private Sub cmdBack_Click()
    '--------------------------------------------------------------
    ' Memuat halaman sebelumnya (bagian dari route dan navigasi)
    '--------------------------------------------------------------
    rs.MovePrevious
    If rs.BOF = True Then rs.MoveFirst
    GetData
End Sub

Private Sub cmdFirst_Click()
    '--------------------------------------------------------------
    ' Memuat interaksi database ke baris pertama (Forward)
    '--------------------------------------------------------------
    rs.MoveFirst
    GetData
End Sub

Private Sub cmdLast_Click()
    '--------------------------------------------------------------
    ' Memuat interaksi database langsung ke baris terakhir
    '--------------------------------------------------------------
    rs.MoveLast
    GetData
End Sub

Private Sub cmdNext_Click()
    '--------------------------------------------------------------
    ' Memuat interaksi Database ke baris selanjutnya (Next)
    '--------------------------------------------------------------
    rs.MoveNext
    If rs.EOF = True Then rs.MoveLast
    GetData
End Sub

Private Sub cmdGo_Click()
    '--------------------------------------------------------------
    ' Function yang berinteraksi dengan database ,
    ' yang digunakan untuk mencari rekaman melalui conditional
    ' Input yang digunakan adalah integer dan string
    ' Jika kedua kondisi terpenuhi  , akan ditampilkan rekaman
    ' jika tidak akan dimuat sesuai dengan int/string yang ada
    '--------------------------------------------------------------
    Dim Record As Integer
    If txtGo.Text = "" Then
        MsgBox "Pergi ke Halaman", vbCritical + vbOKOnly, "Error with Input"
    Else
        If IsNumeric(txtGo.Text) = False Then
            MsgBox "Masukan Rekaman Halaman yang dituju", vbCritical + vbOKOnly, "Error with Input"
        Else
            GetNumRows
            Record = txtGo.Text - 1
            If Record < 0 Then Record = 0
            If Record > NumRows Then Record = NumRows - 1
            rs.MoveFirst
            rs.Move (Record)
            GetData
        End If
    End If
    txtGo.Text = ""
End Sub

Private Sub mnuFOpen_Click()
    '--------------------------------------------------------------
    ' Memuat dialog box yang berinteraksi database ketika diclick
    '--------------------------------------------------------------
    With dlgOpen
        .DialogTitle = "Open Database"
        .CancelError = False
        .Filter = "Database Files (*.dat,*.mdb)|*.dat;*.mdb|"
        .Filter = .Filter + "Access Databases (*.mdb)|*.mdb|"
        .Filter = .Filter + "Dat Files (*.dat)|*.dat|"
        .Filter = .Filter + "All Files (*.*)|*.*"
        .InitDir = App.Path
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
    End With
End Sub

Private Sub mnuFExit_Click()
    '--------------------------------------------------------------
    ' Program diakhiri
    '--------------------------------------------------------------
    Unload Me
    End
End Sub

Private Sub mnuTDelete_Click()
    '--------------------------------------------------------------
    ' Verifikasi apakah ingin mendelete database ini ?
    '--------------------------------------------------------------
    Dim Answer As Integer
    Answer = MsgBox("Apakah anda ingin menghapus rekaman ini ?", vbQuestion + vbYesNo, "Hapus rekaman ?")
    If Answer = 6 Then
        rs.Delete
        rs.MoveLast
        GetData
    Else
        GetData
    End If
End Sub

Private Sub mnuTEdit_Click()
    '--------------------------------------------------------------
    ' mnuTEdit digunakan untuk mengedit text box
    ' fitur ini berinteraksi dengan DoWhat untuk editing rekaman
    '--------------------------------------------------------------
    cmdSave.Enabled = True
    cmdClear.Enabled = True
    txtTitle.Locked = False
    txtSerial.Locked = False
    txtCompany.Locked = False
    cmdCancel.Visible = True
    cmdCancel.Enabled = True
    cmdGo.Visible = False
    cmdGo.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    cmdFirst.Enabled = False
    cmdBack.Enabled = False
    DoWhat = 1
End Sub

Private Sub mnuTNew_Click()
    '--------------------------------------------------------------
    ' mnuTNew_Click akan membuat text box menjadi kosong
    ' fitur ini digunakan untuk (Save , Clear , Cancel)
    ' yang berinteraksi dengan DoWhat 
    '--------------------------------------------------------------
    cmdSave.Enabled = True
    cmdClear.Enabled = True
    cmdGo.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    cmdFirst.Enabled = False
    cmdBack.Enabled = False
    txtTitle.Text = "Title"
    txtTitle.Locked = False
    txtSerial.Text = ""
    txtSerial.Locked = False
    txtCompany.Text = ""
    txtCompany.Locked = False
    cmdCancel.Visible = True
    cmdCancel.Enabled = True
    cmdGo.Visible = False
    DoWhat = 0
End Sub

Private Sub mnuTPrint_Click()
    '--------------------------------------------------------------
    ' Menampilkan database pada MS Access
    '--------------------------------------------------------------
    rptSoftware.Show
End Sub

Private Sub mnuHAbout_Click()
    '--------------------------------------------------------------
    ' Menampilkan nama-nama pembuat program pada about box
    '--------------------------------------------------------------
    Dim Line1, Title
Line1 = "Software Inventory : Riani , Dimas , Purnomo , Heri , Rizki "
    Title = "About Software Inventory"
    MsgBox Line1, vbInformation + vbOKOnly, Title
End Sub

Private Sub mnuHHelp_Click()
    '--------------------------------------------------------------
    ' Code penyanggah jika terdapat error .
    ' Note : Belum di implementasi
    '--------------------------------------------------------------
    Dim nRun
    nRun = Shell("hh.exe " & App.Path & "\help.chm", vbMaximizedFocus)
End Sub

Public Function GetData()
    '--------------------------------------------------------------
    ' Function ini digunakan untuk membaca/memuat database
    '--------------------------------------------------------------
    fraRecord.Caption = "Record:" & rs.Fields("ID")
    txtTitle.Text = rs.Fields("Title")
    txtSerial.Text = rs.Fields("Serial")
    txtCompany.Text = rs.Fields("Company")
End Function

Public Function GetNumRows()
    '--------------------------------------------------------------
    ' Function dibawah ini menyatakan jumlah baris pada Database
    '--------------------------------------------------------------
    NumRows = 0
    rs.MoveFirst
    Do While Not rs.EOF
        NumRows = NumRows + 1
        rs.MoveNext
    Loop
End Function
