VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCountry 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Language"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   ControlBox      =   0   'False
   Icon            =   "frmCountry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Tab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   8438015
      TabCaption(0)   =   "Choose a Language"
      TabPicture(0)   =   "frmCountry.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "btnOK"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "btnCancel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "List1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "List1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Timer1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Change Country Informations"
      TabPicture(1)   =   "frmCountry.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "btnDelete"
      Tab(1).Control(1)=   "btnNew"
      Tab(1).Control(2)=   "List2"
      Tab(1).Control(3)=   "rsCountry"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(6)=   "CMD1"
      Tab(1).ControlCount=   7
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   6000
         Top             =   3360
      End
      Begin Project1.LaVolpeButton btnDelete 
         Height          =   495
         Left            =   -69840
         TabIndex        =   23
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTNICON         =   "frmCountry.frx":047A
         BTYPE           =   3
         TX              =   "&Delete Country"
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
         MICON           =   "frmCountry.frx":05D4
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
         Left            =   -72720
         TabIndex        =   22
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTNICON         =   "frmCountry.frx":05F0
         BTYPE           =   3
         TX              =   "&New Country"
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
         MICON           =   "frmCountry.frx":0CC2
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
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   3735
         Left            =   -74880
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   2955
         Index           =   1
         Left            =   2160
         TabIndex        =   20
         Top             =   840
         Width           =   3495
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   2955
         Index           =   0
         Left            =   1200
         TabIndex        =   19
         Top             =   840
         Width           =   975
      End
      Begin VB.Data rsCountry 
         Appearance      =   0  'Flat
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Master\Recipe\Recipe.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   315
         Left            =   -72600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Country"
         Top             =   240
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Frame Frame1 
         Caption         =   "Flag"
         Height          =   1095
         Left            =   -72720
         TabIndex        =   15
         Top             =   2520
         Width           =   4455
         Begin VB.CommandButton btnCopy 
            Height          =   615
            Left            =   840
            Picture         =   "frmCountry.frx":0CDE
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Copy Picture to Clipboard"
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton btnPaste 
            Height          =   615
            Left            =   120
            Picture         =   "frmCountry.frx":13A0
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Copy picture from clipboard"
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton btnOpen 
            Height          =   615
            Left            =   1560
            Picture         =   "frmCountry.frx":1A62
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Read picture from disk"
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton btnDeletePicture 
            Height          =   615
            Left            =   2280
            Picture         =   "frmCountry.frx":2124
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Delete picture"
            Top             =   360
            Width           =   615
         End
         Begin VB.Image Picture1 
            DataField       =   "CountryFlag"
            DataSource      =   "rsCountry"
            Height          =   735
            Left            =   3360
            Stretch         =   -1  'True
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   -72720
         TabIndex        =   4
         Top             =   360
         Width           =   4455
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CountryFix"
            DataSource      =   "rsCountry"
            Height          =   285
            Left            =   2400
            TabIndex        =   9
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CountryPrefix"
            DataSource      =   "rsCountry"
            Height          =   285
            Left            =   2400
            TabIndex        =   8
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "ExchangeRate"
            DataSource      =   "rsCountry"
            Height          =   285
            Left            =   2400
            TabIndex        =   7
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Currency"
            DataSource      =   "rsCountry"
            Height          =   285
            Left            =   2400
            TabIndex        =   6
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Country"
            DataSource      =   "rsCountry"
            Height          =   285
            Left            =   2400
            TabIndex        =   5
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Country Short:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   14
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Phone Prefix:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Exchange Rate:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Currency:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Country Name:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   2175
         End
      End
      Begin Project1.LaVolpeButton btnCancel 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Cancel"
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Cancel"
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
         MICON           =   "frmCountry.frx":226E
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
      Begin Project1.LaVolpeButton btnOK 
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         ToolTipText     =   "OK, change language"
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&OK"
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
         MICON           =   "frmCountry.frx":228A
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
      Begin MSComDlg.CommonDialog CMD1 
         Left            =   -72840
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Choose one of the folowing Countries:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim bookmarks1() As Variant
Dim bookmarks2() As Variant
Dim boolNewRecord As Boolean
Dim rsLanguage As Recordset
Dim rsMyRec As Recordset
Private Sub LoadList1()
    On Error Resume Next
    List1(0).Clear
    List1(1).Clear
    With rsCountry.Recordset
        .MoveLast
        .MoveFirst
        ReDim bookmarks1(.RecordCount)
        Do While Not .EOF
            List1(0).AddItem .Fields("CountryFix")
            List1(0).ItemData(List1(0).NewIndex) = List1(0).ListCount - 1
            bookmarks1(List1(0).ListCount - 1) = .Bookmark
            List1(1).AddItem .Fields("Country")
        .MoveNext
        Loop
    End With
End Sub


Private Sub LoadList2()
    On Error Resume Next
    List2.Clear
    With rsCountry.Recordset
        .MoveLast
        .MoveFirst
        ReDim bookmarks2(.RecordCount)
        Do While Not .EOF
            List2.AddItem .Fields("Country")
            List2.ItemData(List2.NewIndex) = List2.ListCount - 1
            bookmarks2(List2.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub


Private Sub ShowText()
Dim strHelp As String
    On Error Resume Next
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then
                .Edit
                If IsNull(.Fields("Form")) Then
                    .Fields("Form") = Me.Caption
                Else
                    Me.Caption = .Fields("Form")
                End If
                For i = 0 To 5
                    If IsNull(i + 2) Then
                        .Fields(i + 2) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 2)
                    End If
                Next
                If IsNull(.Fields("Frame1")) Then
                    .Fields("Frame1") = Frame1.Caption
                Else
                    Frame1.Caption = .Fields("Frame1")
                End If
                If IsNull(.Fields("btnCancel")) Then
                    .Fields("btnCancel") = btnCancel.Caption
                Else
                    btnCancel.Caption = .Fields("btnCancel")
                End If
                If IsNull(.Fields("btnOK")) Then
                    .Fields("btnOK") = btnOK.Caption
                Else
                    btnOK.Caption = .Fields("btnOK")
                End If
                If IsNull(.Fields("btnNew")) Then
                    .Fields("btnNew") = btnNew.Caption
                Else
                    btnNew.Caption = .Fields("btnNew")
                End If
                If IsNull(.Fields("btnDelete")) Then
                    .Fields("btnDelete") = btnDelete.Caption
                Else
                    btnDelete.Caption = .Fields("btnDelete")
                End If
                If IsNull(.Fields("btnPaste")) Then
                    .Fields("btnPaste") = btnPaste.ToolTipText
                Else
                    btnPaste.ToolTipText = .Fields("btnPaste")
                End If
                If IsNull(.Fields("btnCopy")) Then
                    .Fields("btnCopy") = btnCopy.ToolTipText
                Else
                    btnCopy.ToolTipText = .Fields("btnCopy")
                End If
                If IsNull(.Fields("btnOpen")) Then
                    .Fields("btnOpen") = btnOpen.ToolTipText
                Else
                    btnOpen.ToolTipText = .Fields("btnOpen")
                End If
                If IsNull(.Fields("btnDeletePicture")) Then
                    .Fields("btnDeletePicture") = btnDeletePicture.ToolTipText
                Else
                    btnDeletePicture.ToolTipText = .Fields("btnDeletePicture")
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("Tab10")) Then
                    .Fields("Tab10") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab10")
                End If
                Tab1.Tab = 1
                If IsNull(.Fields("Tab11")) Then
                    .Fields("Tab11") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab11")
                End If
                Tab1.Tab = 0
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
            
        .MoveFirst
        .AddNew
        .Fields("Language") = m_FileExt
        For i = 0 To 5
        .Fields(i + 2) = Label1(i).Caption
        Next
        .Fields("Frame1") = Frame1.Caption
        .Fields("btnCancel") = btnCancel.Caption
        .Fields("btnOK") = btnOK.Caption
        .Fields("btnNew") = btnNew.Caption
        .Fields("btnDelete") = btnDelete.Caption
        .Fields("btnPaste") = btnPaste.ToolTipText
        .Fields("btnCopy") = btnCopy.ToolTipText
        .Fields("btnOpen") = btnOpen.ToolTipText
        .Fields("btnDeletePicture") = btnDeletePicture.ToolTipText
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 0
        .Update
    End With
End Sub

Private Sub btnCopy_Click()
    On Error Resume Next
    'Clipboard.SetData Picture1.Picture(vbCFDIB)
End Sub

Private Sub btnDelete_Click()
    rsCountry.Recordset.Delete
    LoadList2
End Sub

Private Sub btnDeletePicture_Click()
    On Error Resume Next
    Picture1.Picture = LoadPicture()
End Sub

Private Sub btnNew_Click()
    rsCountry.Recordset.AddNew
    boolNewRecord = True
    Text2.SetFocus
End Sub

Private Sub btnOk_Click()
    On Error GoTo errChangeLanguage
    With rsMyRec
        .Edit
        .Fields("LanguageScreen") = rsCountry.Recordset.Fields("CountryFix")
        .Update
    End With
    
    'check if the language is present in the database
    If IsLanguagePresent(rsLanguage, Trim(rsCountry.Recordset.Fields("CountryFix"))) Then
        m_FileExt = Trim(rsCountry.Recordset.Fields("CountryFix"))
        frmRecipe.ReadText
        Exit Sub
    Else   'we did not have this language in the Recipe.mdb database
        MakeNewLanguage (Trim(rsCountry.Recordset.Fields("CountryFix"))) 'make new recordset for each form in database
    End If
    m_FileExt = Trim(Trim(rsCountry.Recordset.Fields("CountryFix")))
    ShowText    'show this form-text
    frmRecipe.ReadText
    Unload Me
    Exit Sub
    
errChangeLanguage:
    Beep
    MsgBox Err.Description, vbExclamation, "Change Language"
    Err.Clear
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOpen_Click()
    With CMD1
        .FileName = ""
        .DialogTitle = "Load Picture from disk"
        .Filter = "(*.bmp)|*.bmp|(*.pcx)|*.pcx|(*.jpg)|*.jpg"
        .FilterIndex = 1
        .ShowOpen
        Picture1.Picture = LoadPicture(.FileName)
    End With
End Sub

Private Sub btnPaste_Click()
    Picture1.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsCountry.Refresh
    LoadList1
    LoadList2
    Call ShowText
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsCountry.DatabaseName = m_sData
    Set rsMyRec = m_dbData.OpenRecordset("MyRec")
    Set rsLanguage = m_dbData.OpenRecordset("frmCountry")
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsCountry.Recordset.Close
    rsLanguage.Close
    rsMyRec.Close
    Erase bookmarks1
    Erase bookmarks2
    frmRecipe.LoadCountry
    Set frmCountry = Nothing
End Sub

Private Sub List1_Click(Index As Integer)
    On Error Resume Next
    i = List1(Index).ListIndex
    List1(0).ListIndex = i
    List1(1).ListIndex = i
    rsCountry.Recordset.Bookmark = bookmarks1(List1(0).ItemData(List1(0).ListIndex))
End Sub

Private Sub List2_Click()
    On Error Resume Next
    rsCountry.Recordset.Bookmark = bookmarks2(List2.ItemData(List2.ListIndex))
End Sub

Private Sub Text2_LostFocus()
    If boolNewRecord Then
        On Error GoTo errText2_Click
        With rsCountry.Recordset
            .Fields("Country") = Text2.Text
            .Update
            .Bookmark = rsCountry.Recordset.LastModified
            LoadList2
        End With
        boolNewRecord = False
        Text3.SetFocus
    End If
    Exit Sub
    
errText2_Click:
    Beep
    MsgBox Err.Description, vbCritical, "New Record"
    Err.Clear
    boolNewRecord = False
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    List1(1).TopIndex = List1(0).TopIndex
End Sub


