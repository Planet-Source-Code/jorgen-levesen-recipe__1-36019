VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmScreenLanguage 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Screen Text"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.LaVolpeButton btnExit 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   7080
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   661
      BTNICON         =   "frmScreenLanguage.frx":0000
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
      MICON           =   "frmScreenLanguage.frx":015A
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
   Begin TabDlg.SSTab Tab1 
      Height          =   6855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   11
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   8438015
      TabCaption(0)   =   "Country"
      TabPicture(0)   =   "frmScreenLanguage.frx":0176
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DBGrid1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Data1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Dimension"
      TabPicture(1)   =   "frmScreenLanguage.frx":0192
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1(1)"
      Tab(1).Control(1)=   "Data2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Recipe"
      TabPicture(2)   =   "frmScreenLanguage.frx":01AE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DBGrid1(2)"
      Tab(2).Control(1)=   "Data3"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Recipe Types"
      TabPicture(3)   =   "frmScreenLanguage.frx":01CA
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DBGrid1(3)"
      Tab(3).Control(1)=   "Data4"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Ingredients"
      TabPicture(4)   =   "frmScreenLanguage.frx":01E6
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "DBGrid1(4)"
      Tab(4).Control(1)=   "Data5"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Screen Text"
      TabPicture(5)   =   "frmScreenLanguage.frx":0202
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "DBGrid1(5)"
      Tab(5).Control(1)=   "Data6"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Search for Ingredients"
      TabPicture(6)   =   "frmScreenLanguage.frx":021E
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "DBGrid1(6)"
      Tab(6).Control(1)=   "Data7"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Mail"
      TabPicture(7)   =   "frmScreenLanguage.frx":023A
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "DBGrid1(7)"
      Tab(7).Control(1)=   "Data8"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Internet Links"
      TabPicture(8)   =   "frmScreenLanguage.frx":0256
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Data9"
      Tab(8).Control(1)=   "DBGrid1(8)"
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "User Information"
      TabPicture(9)   =   "frmScreenLanguage.frx":0272
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Data10"
      Tab(9).Control(1)=   "DBGrid1(9)"
      Tab(9).ControlCount=   2
      TabCaption(10)  =   "Write To Me"
      TabPicture(10)  =   "frmScreenLanguage.frx":028E
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "DBGrid1(10)"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).Control(1)=   "Data11"
      Tab(10).Control(1).Enabled=   0   'False
      Tab(10).ControlCount=   2
      Begin VB.Data Data11 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Master\Recipe\Recipe.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -73680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmWriteToMe"
         Top             =   2520
         Width           =   1140
      End
      Begin VB.Data Data10 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Master\Recipe\Recipe.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -73680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmUser"
         Top             =   2520
         Width           =   1140
      End
      Begin VB.Data Data9 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Master\Recipe\Recipe.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -73680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmLinks"
         Top             =   2520
         Width           =   1140
      End
      Begin VB.Data Data8 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Master\Recipe\Recipe.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -73680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmMail"
         Top             =   2700
         Width           =   1140
      End
      Begin VB.Data Data7 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Master\Recipe\Recipe.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -73680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmHaveIngredients"
         Top             =   2820
         Width           =   1140
      End
      Begin VB.Data Data6 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Master\Recipe\Recipe.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -73680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmScreenLanguage"
         Top             =   2460
         Width           =   1140
      End
      Begin VB.Data Data5 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Recipes\Recipe.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -73560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmRecipeTypes"
         Top             =   2640
         Width           =   1140
      End
      Begin VB.Data Data4 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Recipes\Recipe.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -73560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmRecipeTypes"
         Top             =   2640
         Width           =   1140
      End
      Begin VB.Data Data3 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Recipes\Recipe.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -73680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmRecipe"
         Top             =   2640
         Width           =   1140
      End
      Begin VB.Data Data2 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Recipes\Recipe.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -73560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmDimension"
         Top             =   2520
         Width           =   1140
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Recipes\Recipe.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   1440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmCountry"
         Top             =   2520
         Width           =   1140
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenLanguage.frx":02AA
         Height          =   5415
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "frmScreenLanguage.frx":02BE
         TabIndex        =   1
         Top             =   1320
         Width           =   8055
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenLanguage.frx":0C94
         Height          =   5415
         Index           =   1
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":0CA8
         TabIndex        =   2
         Top             =   1260
         Width           =   8055
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenLanguage.frx":167E
         Height          =   5415
         Index           =   2
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":1692
         TabIndex        =   3
         Top             =   1260
         Width           =   8055
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenLanguage.frx":2068
         Height          =   5415
         Index           =   3
         Left            =   -74760
         OleObjectBlob   =   "frmScreenLanguage.frx":207C
         TabIndex        =   4
         Top             =   1260
         Width           =   7935
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenLanguage.frx":2A52
         Height          =   5415
         Index           =   4
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":2A66
         TabIndex        =   6
         Top             =   1260
         Width           =   8055
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenLanguage.frx":343C
         Height          =   5415
         Index           =   5
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":3450
         TabIndex        =   7
         Top             =   1260
         Width           =   8055
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenLanguage.frx":3E26
         Height          =   5415
         Index           =   6
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":3E3A
         TabIndex        =   8
         Top             =   1320
         Width           =   8055
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenLanguage.frx":4810
         Height          =   5415
         Index           =   7
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":4824
         TabIndex        =   9
         Top             =   1320
         Width           =   8055
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenLanguage.frx":51FA
         Height          =   5415
         Index           =   8
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":520E
         TabIndex        =   10
         Top             =   1320
         Width           =   8055
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenLanguage.frx":5BE4
         Height          =   5415
         Index           =   9
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":5BF9
         TabIndex        =   11
         Top             =   1320
         Width           =   8055
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenLanguage.frx":65CF
         Height          =   5415
         Index           =   10
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":65E4
         TabIndex        =   12
         Top             =   1320
         Width           =   8055
      End
   End
End
Attribute VB_Name = "frmScreenLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLanguage As Recordset
Private Sub ReadText()
    'find YOUR sLanguage text
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
                Tab1.Tab = 2
                If IsNull(.Fields("Tab12")) Then
                    .Fields("Tab12") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab12")
                End If
                Tab1.Tab = 3
                If IsNull(.Fields("Tab13")) Then
                    .Fields("Tab13") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab13")
                End If
                Tab1.Tab = 4
                If IsNull(.Fields("Tab14")) Then
                    .Fields("Tab14") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab14")
                End If
                Tab1.Tab = 5
                If IsNull(.Fields("Tab15")) Then
                    .Fields("Tab15") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab15")
                End If
                Tab1.Tab = 6
                If IsNull(.Fields("Tab16")) Then
                    .Fields("Tab16") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab16")
                End If
                Tab1.Tab = 7
                If IsNull(.Fields("Tab17")) Then
                    .Fields("Tab17") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab17")
                End If
                Tab1.Tab = 8
                If IsNull(.Fields("Tab18")) Then
                    .Fields("Tab18") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab18")
                End If
                Tab1.Tab = 9
                If IsNull(.Fields("Tab19")) Then
                    .Fields("Tab19") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab19")
                End If
                Tab1.Tab = 10
                If IsNull(.Fields("Tab110")) Then
                    .Fields("Tab110") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab110")
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("btnExit")) Then
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
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 2
        .Fields("Tab12") = Tab1.Caption
        Tab1.Tab = 3
        .Fields("Tab13") = Tab1.Caption
        Tab1.Tab = 4
        .Fields("Tab14") = Tab1.Caption
        Tab1.Tab = 5
        .Fields("Tab15") = Tab1.Caption
        Tab1.Tab = 6
        .Fields("Tab16") = Tab1.Caption
        Tab1.Tab = 7
        .Fields("Tab17") = Tab1.Caption
        Tab1.Tab = 8
        .Fields("Tab18") = Tab1.Caption
        Tab1.Tab = 9
        .Fields("Tab19") = Tab1.Caption
        Tab1.Tab = 10
        .Fields("Tab110") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("btnExit") = btnExit.Caption
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Data1.Refresh
    Data2.Refresh
    Data3.Refresh
    Data4.Refresh
    Data5.Refresh
    Data6.Refresh
    Data7.Refresh
    Data8.Refresh
    Data9.Refresh
    Data10.Refresh
    Data11.Refresh
    ReadText
End Sub

Private Sub Form_Load()
    On Error Resume Next    'it is only text
    Data1.DatabaseName = m_sData
    Data2.DatabaseName = m_sData
    Data3.DatabaseName = m_sData
    Data4.DatabaseName = m_sData
    Data5.DatabaseName = m_sData
    Data6.DatabaseName = m_sData
    Data7.DatabaseName = m_sData
    Data8.DatabaseName = m_sData
    Data9.DatabaseName = m_sData
    Data10.DatabaseName = m_sData
    Data11.DatabaseName = m_sData
    Set rsLanguage = m_dbData.OpenRecordset("frmScreenLanguage")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Data1.Recordset.Close
    Data2.Recordset.Close
    Data3.Recordset.Close
    Data4.Recordset.Close
    Data5.Recordset.Close
    Data6.Recordset.Close
    Data7.Recordset.Close
    Data8.Recordset.Close
    Data9.Recordset.Close
    Data10.Recordset.Close
    Data11.Recordset.Close
    rsLanguage.Close
    Set frmScreenLanguage = Nothing
End Sub
