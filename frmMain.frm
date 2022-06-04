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
   MaxButton       =   0   'False
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
      Width           =   3495
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
' Declare all global variables here.  Also, to make sure each
' variable is declared, set the Option Explicit.
'--------------------------------------------------------------
Option Explicit
    Dim db As Database
    Dim rs As Recordset
    Dim DoWhat As Integer
    Dim NumRows As Integer

Private Sub Form_Load()
    '--------------------------------------------------------------
    ' Sets up the Database and Recordset objects.  Also sets the
    ' program's state to View mode and populates the text boxes
    ' with the first record in the database.
    '--------------------------------------------------------------
    Set db = OpenDatabase(App.Path & "\data.mdb")
    Set rs = db.OpenRecordset("tblSoftware", dbOpenDynaset)
    cmdSave.Enabled = False
    cmdClear.Enabled = False
    cmdCancel.Enabled = False
    cmdCancel.Visible = False
    txtTitle.Text = ""
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
    ' Saves the new or edited record to the database.  The DoWhat
    ' variable is used to determine if the database edit will be an
    ' Update statement or an Add statement.  It also resets the
    ' program's state to View Mode.
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
    ' Depending on the DoWhat variable, it either clears the text
    ' boxes (New Entry) or resets the data (Edit Entry)
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
    ' Only visible in Editing Mode.  It cancels the edit and
    ' returns the program to the View Mode.
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
    ' Procedes to the previous record and displays it.
    '--------------------------------------------------------------
    rs.MovePrevious
    If rs.BOF = True Then rs.MoveFirst
    GetData
End Sub

Private Sub cmdFirst_Click()
    '--------------------------------------------------------------
    ' Procedes to the first record and displays it.
    '--------------------------------------------------------------
    rs.MoveFirst
    GetData
End Sub

Private Sub cmdLast_Click()
    '--------------------------------------------------------------
    ' Procedes to the last record and displays it.
    '--------------------------------------------------------------
    rs.MoveLast
    GetData
End Sub

Private Sub cmdNext_Click()
    '--------------------------------------------------------------
    ' Procedes to the next record and displays it.
    '--------------------------------------------------------------
    rs.MoveNext
    If rs.EOF = True Then rs.MoveLast
    GetData
End Sub

Private Sub cmdGo_Click()
    '--------------------------------------------------------------
    ' This function checks the input for being an integer and
    ' non-empty.  If both conditions are met, it checks to see that
    ' the user-inputed number is located in the database.  If so,
    ' it displayes the given record.  If not, it displays either
    ' the first or last record depending on the number inputed.
    '--------------------------------------------------------------
    Dim Record As Integer
    If txtGo.Text = "" Then
        MsgBox "Enter an integer only", vbCritical + vbOKOnly, "Error with Input"
    Else
        If IsNumeric(txtGo.Text) = False Then
            MsgBox "Enter an integer only", vbCritical + vbOKOnly, "Error with Input"
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
    ' Displays the common dialog box to open a new database.
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
    ' Exits the program
    '--------------------------------------------------------------
    Unload Me
    End
End Sub

Private Sub mnuTDelete_Click()
    '--------------------------------------------------------------
    ' First, it verifies that the user wants to delete the record
    ' and based on the Yes/No answer, it deletes the record.
    '--------------------------------------------------------------
    Dim Answer As Integer
    Answer = MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Verify Deletion of Record")
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
    ' All this does is clear the text boxes, enables them for
    ' editing, forces the navigation buttons to be disabled and
    ' enables the editing buttons (Save, Clear, Cancel)
    ' It also sets the "DoWhat" setting, used when saving/updating
    ' the database.
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
    ' All this does is clear the text boxes, enables them for
    ' editing, forces the navigation buttons to be disabled and
    ' enables the editing buttons (Save, Clear, Cancel)
    ' It also sets the "DoWhat" setting, used when saving/updating
    ' the database.
    '--------------------------------------------------------------
    cmdSave.Enabled = True
    cmdClear.Enabled = True
    cmdGo.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    cmdFirst.Enabled = False
    cmdBack.Enabled = False
    txtTitle.Text = ""
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
    ' Displays the Report of all the software information in the
    ' database and displays it in a standard Access report format.
    '--------------------------------------------------------------
    rptSoftware.Show
End Sub

Private Sub mnuHAbout_Click()
    '--------------------------------------------------------------
    ' Displayes a simple "about box" in the form of a message box.
    '--------------------------------------------------------------
    Dim Line1, Title
    Line1 = "Software Inventory\n Rizki F "
    Title = "About Software Inventory"
    MsgBox Line1, vbInformation + vbOKOnly, Title
End Sub

Private Sub mnuHHelp_Click()
    '--------------------------------------------------------------
    ' Not implimented.  I haven't written the help file yet.
    ' However I do know this code works, as I use it in another
    ' working application.
    '--------------------------------------------------------------
    Dim nRun
    nRun = Shell("hh.exe " & App.Path & "\help.chm", vbMaximizedFocus)
End Sub

Public Function GetData()
    '--------------------------------------------------------------
    ' This is responsible for reading the data out of the database.
    '--------------------------------------------------------------
    fraRecord.Caption = "Record:" & rs.Fields("ID")
    txtTitle.Text = rs.Fields("Title")
    txtSerial.Text = rs.Fields("Serial")
    txtCompany.Text = rs.Fields("Company")
End Function

Public Function GetNumRows()
    '--------------------------------------------------------------
    ' All this function does is count the number of rows in the
    ' database.
    '--------------------------------------------------------------
    NumRows = 0
    rs.MoveFirst
    Do While Not rs.EOF
        NumRows = NumRows + 1
        rs.MoveNext
    Loop
End Function
