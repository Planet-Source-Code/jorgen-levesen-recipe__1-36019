VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAbout 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab Tab1 
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Description"
      TabPicture(0)   =   "frmAbout.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblVersion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTitle"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdOK"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "picIcon"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Credits"
      TabPicture(1)   =   "frmAbout.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3(2)"
      Tab(1).Control(1)=   "Label3(1)"
      Tab(1).Control(2)=   "Label3(0)"
      Tab(1).ControlCount=   3
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   480
         Picture         =   "frmAbout.frx":047A
         ScaleHeight     =   480
         ScaleMode       =   0  'User
         ScaleWidth      =   480
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   600
         Width           =   510
      End
      Begin Project1.LaVolpeButton cmdOK 
         Height          =   375
         Left            =   5040
         TabIndex        =   1
         ToolTipText     =   "Exit"
         Top             =   5160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTNICON         =   "frmAbout.frx":08BC
         BTYPE           =   3
         TX              =   "&Exit"
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
         MICON           =   "frmAbout.frx":0A16
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":0A32
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   0
         Left            =   -74760
         TabIndex        =   9
         Top             =   1080
         Width           =   5535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Thanks to John McGlothlin for the ccOutlookSendMail Class"
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
         Index           =   1
         Left            =   -74760
         TabIndex        =   8
         Top             =   2880
         Width           =   5655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Thanks to LaVolpe for the XP-like buttons"
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
         Index           =   2
         Left            =   -74760
         TabIndex        =   7
         Top             =   3240
         Width           =   5775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "See My Web page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   5040
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2295
         Left            =   480
         TabIndex        =   5
         Top             =   2040
         Width           =   5415
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Application Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1200
         TabIndex        =   4
         Tag             =   "Application Title"
         Top             =   600
         Width           =   4725
      End
      Begin VB.Label lblVersion 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Version: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Tag             =   "Version"
         Top             =   840
         Width           =   4725
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lColours() As Long
Dim lPos() As Long
Dim rsLanguage As Recordset
Dim iLoop As Long
Private Sub ReadText()
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
                If IsNull(.Fields("Label1")) Then
                    .Fields("Label1") = Label1.Caption
                Else
                    Label1.Caption = .Fields("Label1")
                End If
                If IsNull(.Fields("Label2")) Then
                    .Fields("Label2") = Label2.Caption
                Else
                    Label2.Caption = .Fields("Label2")
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("Tab1(0)")) Then
                    .Fields("Tab1(0)") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab1(0)")
                End If
                Tab1.Tab = 1
                If IsNull(.Fields("Tab1(1)")) Then
                    .Fields("Tab1(1)") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab1(1)")
                End If
                Tab1.Tab = 0
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        .Fields("Label1") = Label1.Caption
        .Fields("Label2") = Label2.Caption
        .Fields("Tab1(0)") = "Description"
        .Fields("Tab1(1)") = "Credits"
        .Update
    End With
End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Call ReadText
    cmdOK.RefreshButton
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Show
    Me.Move 0, 0
    Set rsLanguage = m_dbData.OpenRecordset("frmAbout")
    lblVersion.Caption = lblVersion.Caption & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub
Private Sub cmdOK_Click()
        Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLanguage.Close
    Set frmAbout = Nothing
End Sub

Private Sub Label2_Click()
Dim X As Long
    X = Shell("explorer http://www.levesen.com")
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = &H80FF80
End Sub

Private Sub Tab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = &H0&
End Sub
