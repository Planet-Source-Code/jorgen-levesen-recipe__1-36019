VERSION 5.00
Begin VB.Form frmLinks 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internet Recipe Links"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin Project1.LaVolpeButton btnConnect 
         Height          =   495
         Left            =   9360
         TabIndex        =   20
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTNICON         =   "frmLinks.frx":0000
         BTYPE           =   3
         TX              =   "&Connect"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16761024
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmLinks.frx":031A
         ALIGN           =   1
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   3
         SHOWF           =   -1  'True
         BSTYLE          =   0
         OPTVAL          =   0   'False
         OPTMOD          =   0   'False
         GStart          =   0
         GStop           =   16711680
         GStyle          =   0
      End
      Begin Project1.LaVolpeButton btnDelete 
         Height          =   495
         Left            =   6600
         TabIndex        =   19
         Top             =   6120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTNICON         =   "frmLinks.frx":0336
         BTYPE           =   3
         TX              =   "&Delete"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16761024
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmLinks.frx":0490
         ALIGN           =   1
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
         OPTVAL          =   0   'False
         OPTMOD          =   0   'False
         GStart          =   0
         GStop           =   16711680
         GStyle          =   0
      End
      Begin Project1.LaVolpeButton btnNew 
         Height          =   495
         Left            =   4920
         TabIndex        =   18
         Top             =   6120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTNICON         =   "frmLinks.frx":04AC
         BTYPE           =   3
         TX              =   "&New Link"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16761024
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmLinks.frx":0B7E
         ALIGN           =   1
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
         OPTVAL          =   0   'False
         OPTMOD          =   0   'False
         GStart          =   0
         GStop           =   16711680
         GStyle          =   0
      End
      Begin Project1.LaVolpeButton btnExit 
         Height          =   495
         Left            =   9360
         TabIndex        =   17
         Top             =   6120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTNICON         =   "frmLinks.frx":0B9A
         BTYPE           =   3
         TX              =   "&Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16761024
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmLinks.frx":0CF4
         ALIGN           =   1
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
         OPTVAL          =   0   'False
         OPTMOD          =   0   'False
         GStart          =   0
         GStop           =   16711680
         GStyle          =   0
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LinkLastUsed"
         DataSource      =   "rsLinks"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   9240
         MaxLength       =   50
         TabIndex        =   16
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LinkInDatabase"
         DataSource      =   "rsLinks"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   9240
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LinkPassWord"
         DataSource      =   "rsLinks"
         Height          =   285
         Index           =   4
         Left            =   4920
         MaxLength       =   50
         TabIndex        =   12
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LinkUserName"
         DataSource      =   "rsLinks"
         Height          =   285
         Index           =   3
         Left            =   4920
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1800
         Width           =   3975
      End
      Begin VB.Data rsLinks 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\MasterPlan\Schedules.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   7800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Links"
         Top             =   120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LinkNote"
         DataSource      =   "rsLinks"
         Height          =   3045
         Index           =   2
         Left            =   4920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3000
         Width           =   6015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LinkHyper"
         DataSource      =   "rsLinks"
         Height          =   285
         Index           =   1
         Left            =   4920
         MaxLength       =   70
         TabIndex        =   6
         Top             =   1200
         Width           =   6015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LinkName"
         DataSource      =   "rsLinks"
         Height          =   285
         Index           =   0
         Left            =   4920
         TabIndex        =   3
         Top             =   600
         Width           =   4335
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   6075
         Left            =   360
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Date Used::"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   9240
         TabIndex        =   15
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Stored In Database:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   9240
         TabIndex        =   14
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Link User Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   4920
         TabIndex        =   10
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Link Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   9
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Link Note:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   7
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Link URL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   5
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Link Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Link Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vLinkBook() As Variant
Dim bNewRecord As Boolean
Dim rsLanguage As Recordset
Public Sub SelectRecords()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM Links"
    rsLinks.RecordSource = Sql
    rsLinks.Refresh
End Sub

Private Sub ReadText()
Dim sHelp As String
    'On Error Resume Next    'this is only text
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then
                .Edit
                If IsNull("Form") Then
                    .Fields("Form") = Me.Caption
                Else
                    Me.Caption = .Fields("Form")
                End If
                For i = 0 To 7
                    If IsNull(i + 2) Then
                        .Fields(i + 2) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 2)
                    End If
                Next
                If IsNull("btnNew") Then
                    .Fields("btnNew") = btnNew.Caption
                Else
                    btnNew.Caption = .Fields("btnNew")
                End If
                If IsNull("btnDelete") Then
                    .Fields("btnDelete") = btnDelete.Caption
                Else
                    btnDelete.Caption = .Fields("btnDelete")
                End If
                If IsNull("btnConnect") Then
                    .Fields("btnConnect") = btnConnect.Caption
                Else
                    btnConnect.Caption = .Fields("btnConnect")
                End If
                If IsNull("btnExit") Then
                    .Fields("btnExit") = btnExit.Caption
                Else
                    btnExit.Caption = .Fields("btnExit")
                End If
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        For i = 0 To 7
            .Fields(i + 2) = Label1(i).Caption
        Next
        .Fields("btnNew") = btnNew.Caption
        .Fields("btnDelete") = btnDelete.Caption
        .Fields("btnConnect") = btnConnect.Caption
        .Fields("btnExit") = btnExit.Caption
        .Update
    End With
End Sub

Private Sub LoadList1()
    On Error Resume Next
    List1.Clear
    With rsLinks.Recordset
        .MoveLast
        .MoveFirst
        ReDim vLinkBook(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("LinkName")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            vLinkBook(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Public Sub UpdateLinkDate()
    On Error Resume Next
    With rsLinks.Recordset
        .Edit
        .Fields("LinkLastUsed") = Format(Now, "dd.mm.yyyy")
        .Update
        .Bookmark = .LastModified
    End With
End Sub

Private Sub btnConnect_Click()
Dim iret As Long, strLink As String
    On Error GoTo errConnectLink
        If Len(Text1(1).Text) <> 0 Then
        strLink = Text1(1).Text
        strLink = Left$(strLink, 7)
        If strLink = "http://" Then
            strLink = Text1(1).Text
        Else
            strLink = "http://" & Text1(1).Text
        End If
        iret = ShellExecute(Me.hWnd, _
                vbNullString, _
                strLink, vbNullString, "c:\", _
                SW_SHOWNORMAL)
        End If
    Exit Sub
    
errConnectLink:
    Beep
    MsgBox Err.Description, vbExclamation, "Internet"
    Err.Clear
End Sub

Private Sub btnDelete_Click()
    On Error GoTo errDelete
    rsLinks.Recordset.Delete
    LoadList1
    List1.ListIndex = 0
    Exit Sub
    
errDelete:
    Beep
    MsgBox Err.Description, vbCritical, "Delete a link"
    Err.Clear
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnNew_Click()
    bNewRecord = True
    rsLinks.Recordset.AddNew
    Text1(0).SetFocus
End Sub

Private Sub Form_Activate()
    'On Error Resume Next
    rsLinks.Refresh
    ReadText
    LoadList1
    List1.ListIndex = 0
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsLinks.DatabaseName = m_sData
    Set rsLanguage = m_dbData.OpenRecordset("frmLinks")
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLinks.Recordset.Close
    rsLanguage.Close
    Erase vLinkBook
    Set frmLinks = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    rsLinks.Recordset.Bookmark = vLinkBook(List1.ItemData(List1.ListIndex))
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo errNewLink
    Select Case Index
    Case 0
        If bNewRecord Then
            With rsLinks.Recordset
                .Fields("LinkName") = Trim(Text1(0).Text)
                .Fields("LinkInDatabase") = Format(Now, "dd.mm.yyyy")
                .Update
                LoadList1
                .Bookmark = .LastModified
                bNewRecord = False
            End With
        End If
    Case Else
    End Select
    Exit Sub
    
errNewLink:
    Beep
    MsgBox Err.Description, vbCritical, "New Internet Link"
    Err.Clear
End Sub


