VERSION 5.00
Begin VB.Form frmHaveIngredients 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wish List"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   4560
      Top             =   8280
   End
   Begin VB.Data rsTemp 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\Recipes\Recipe.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SearchRecipe"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data rsRecipe 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Master\Recipe\Recipe.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Recipe"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin Project1.LaVolpeButton btnShow 
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   8280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTNICON         =   "frmHaveIngredients.frx":0000
      BTYPE           =   3
      TX              =   "&Show Recipe"
      ENAB            =   0   'False
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
      MICON           =   "frmHaveIngredients.frx":2842
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
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   8280
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BTNICON         =   "frmHaveIngredients.frx":285E
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
      MICON           =   "frmHaveIngredients.frx":29B8
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
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Recipe suggestions (if any):"
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
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   7935
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   5160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   600
         Width           =   2655
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   3150
         Index           =   1
         Left            =   4200
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   3150
         Index           =   0
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Click to view the ingredients"
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ingredients"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   5160
         TabIndex        =   19
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Serves"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Recipe Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "I would like to make:"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7935
      Begin VB.ComboBox cboRecipeType 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   5295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   " I have the following Ingredients:"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   7935
      Begin VB.Frame Frame4 
         BackColor       =   &H0080C0FF&
         Height          =   2535
         Left            =   6120
         TabIndex        =   16
         Top             =   120
         Width           =   1695
         Begin VB.OptionButton Option1 
            BackColor       =   &H0080C0FF&
            Caption         =   "Just the first Ingredient have to be included"
            ForeColor       =   &H00000000&
            Height          =   1215
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   1200
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H0080C0FF&
            Caption         =   "All Ingredients have to be included"
            ForeColor       =   &H00000000&
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin Project1.LaVolpeButton btnCalculate 
         Height          =   975
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1720
         BTNICON         =   "frmHaveIngredients.frx":29D4
         BTYPE           =   3
         TX              =   "Search for &Recipés"
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
         MICON           =   "frmHaveIngredients.frx":46DE
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
      Begin Project1.LaVolpeButton btnClear 
         Height          =   375
         Left            =   3360
         TabIndex        =   11
         Top             =   2280
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         BTNICON         =   "frmHaveIngredients.frx":46FA
         BTYPE           =   3
         TX              =   "&Clear"
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
         MICON           =   "frmHaveIngredients.frx":4854
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
      Begin Project1.LaVolpeButton btnTransfer 
         Height          =   495
         Left            =   2520
         TabIndex        =   9
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         BTNICON         =   "frmHaveIngredients.frx":4870
         BTYPE           =   3
         TX              =   ""
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
         BCOL            =   12648447
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmHaveIngredients.frx":4CC2
         ALIGN           =   1
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   3
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
         BackColor       =   &H00FFFFFF&
         Height          =   1980
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingredient:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmHaveIngredients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bookmarks1() As Variant
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
                For i = 0 To 3
                    If IsNull(.Fields(i + 2)) Then
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
                If IsNull(.Fields("Frame2")) Then
                    .Fields("Frame2") = Frame2.Caption
                Else
                    Frame2.Caption = .Fields("Frame2")
                End If
                If IsNull(.Fields("Frame3")) Then
                    .Fields("Frame3") = Frame3.Caption
                Else
                    Frame3.Caption = .Fields("Frame3")
                End If
                If IsNull(.Fields("Option1(0)")) Then
                    .Fields("Option1(0)") = Option1(0).Caption
                Else
                    Option1(0).Caption = .Fields("Option1(0)")
                End If
                If IsNull(.Fields("Option1(1)")) Then
                    .Fields("Option1(1)") = Option1(1).Caption
                Else
                    Option1(1).Caption = .Fields("Option1(1)")
                End If
                If IsNull(.Fields("btnTransfer")) Then
                    .Fields("btnTransfer") = btnTransfer.ToolTipText
                Else
                    btnTransfer.ToolTipText = .Fields("btnTransfer")
                End If
                If IsNull(.Fields("btnCalculate")) Then
                    .Fields("btnCalculate") = btnCalculate.Caption
                Else
                    btnCalculate.Caption = .Fields("btnCalculate")
                End If
                If IsNull(.Fields("btnClear")) Then
                    .Fields("btnClear") = btnClear.Caption
                Else
                    btnClear.Caption = .Fields("btnClear")
                End If
                If IsNull(.Fields("btnShow")) Then
                    .Fields("btnShow") = btnShow.Caption
                Else
                    btnShow.Caption = .Fields("btnShow")
                End If
                If IsNull(.Fields("btnExit")) Then
                    .Fields("btnExit") = btnExit.Caption
                Else
                    btnExit.Caption = .Fields("btnExit")
                End If
                
                If IsNull(.Fields("List1")) Then
                    .Fields("List1") = List1(0).ToolTipText
                Else
                    List1(0).ToolTipText = .Fields("List1")
                    List1(1).ToolTipText = .Fields("List1")
                End If
                Exit Sub
             End If
        .MoveNext
        Loop
                
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        For i = 0 To 3
            .Fields(i + 2) = Label1(i).Caption
        Next
        .Fields("Frame1") = Frame1.Caption
        .Fields("Frame2") = Frame2.Caption
        .Fields("Frame3") = Frame3.Caption
        .Fields("Option1(0)") = Option1(0).Caption
        .Fields("Option1(1)") = Option1(1).Caption
        .Fields("btnTransfer") = btnTransfer.ToolTipText
        .Fields("btnCalculate") = btnCalculate.Caption
        .Fields("btnClear") = btnClear.Caption
        .Fields("btnShow") = btnShow.Caption
        .Fields("btnExit") = btnExit.Caption
        .Fields("List1") = List1(0).ToolTipText
        .Update
    End With
End Sub

Private Sub LoadList1()
    On Error Resume Next
    List1(0).Clear
    List1(1).Clear
    With rsTemp.Recordset
        .MoveLast
        .MoveFirst
        ReDim bookmarks1(.RecordCount)
        Do While Not .EOF
            List1(0).AddItem .Fields("Name")
            List1(1).AddItem CInt(.Fields("Serves"))
            List1(0).ItemData(List1(0).NewIndex) = List1(0).ListCount - 1
            bookmarks1(List1(0).ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub
Private Function SearchForRecipe() As Boolean
Dim Sql As String, boolFound As Boolean
    If Len(cboRecipeType.Text) = 0 Then 'have to have one recipe type to search for !
        SearchForRecipe = False
        cboRecipeType.SetFocus
        Exit Function
    End If
    
    'Delete all old records
    On Error Resume Next    'in case recordset is empty
    With rsTemp.Recordset
        .MoveFirst
        Do While Not .EOF
            .Delete
        .MoveNext
        Loop
    End With
    
    On Error GoTo errSelectRecipe
    Sql = "SELECT * FROM Recipe WHERE RecipeType ="
    Sql = Sql & Chr(34) & cboRecipeType.Text & Chr(34)
    Sql = Sql & " ORDER BY Name"
    rsRecipe.RecordSource = Sql
    rsRecipe.Refresh
    
    
    With rsRecipe.Recordset
        If Option1(0).Value = True Then 'all ingredients have to be present
            .MoveFirst
            Do While Not .EOF
                boolFound = True
                For i = 0 To List2.ListCount - 1
                    If InStr(.Fields("Ingredients"), CStr(List2.List(List2.ListIndex))) > 0 Then
                    Else
                        boolFound = False
                    End If
                Next
                If boolFound Then
                    rsTemp.Recordset.AddNew
                    rsTemp.Recordset.Fields("Name") = .Fields("Name")
                    rsTemp.Recordset.Fields("Ingredients") = .Fields("Ingredients")
                    rsTemp.Recordset.Fields("Serves") = .Fields("Serves")
                    rsTemp.Recordset.Update
                End If
            .MoveNext
            Loop
        Else    'just the Text1 ingredient have to be present
            .MoveFirst
            Do While Not .EOF
                If InStr(.Fields("Ingredients"), CStr(Text1.Text)) > 0 Then
                    rsTemp.Recordset.AddNew
                    rsTemp.Recordset.Fields("Name") = .Fields("Name")
                    rsTemp.Recordset.Fields("Ingredients") = .Fields("Ingredients")
                    rsTemp.Recordset.Fields("Serves") = .Fields("Serves")
                    rsTemp.Recordset.Update
                End If
            .MoveNext
            Loop
        End If
    End With
    
    Sql = "SELECT * FROM SearchRecipe ORDER BY Name"
    rsTemp.RecordSource = Sql
    rsTemp.Refresh
    SearchForRecipe = True
    Exit Function
    
errSelectRecipe:
    Err.Clear
    SearchForRecipe = False
End Function

Private Sub btnCalculate_Click()
    List1(0).Clear
    List1(1).Clear
    Text2.Text = ""
    btnShow.Enabled = False
    If SearchForRecipe Then
        btnShow.Enabled = True
        LoadList1
    End If
End Sub

Private Sub btnClear_Click()
    List2.Clear
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnShow_Click()
Dim Sql As String
    On Error Resume Next    'no harm done
    Sql = "SELECT * FROM Recipe WHERE Name ="
    Sql = Sql & Chr(34) & List1(0).List(List1(0).ListIndex) & Chr(34)
    frmRecipe.rsRecipe.RecordSource = Sql
    frmRecipe.rsRecipe.Refresh
    Unload Me
End Sub

Private Sub btnTransfer_Click()
    List2.AddItem Trim(Text1.Text)
    Text1.Text = ""
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsRecipe.Refresh
    rsTemp.Refresh
    ReadText
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsTemp.DatabaseName = m_sData
    rsRecipe.DatabaseName = m_sData
    Set rsLanguage = m_dbData.OpenRecordset("frmHaveIngredients")
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsRecipe.Recordset.Close
    rsTemp.Recordset.Close
    rsLanguage.Close
    Erase bookmarks1
    DBEngine.Idle dbFreeLocks
    Set frmHaveIngredients = Nothing
End Sub

Private Sub List1_Click(Index As Integer)
    On Error Resume Next
    rsTemp.Recordset.Bookmark = bookmarks1(List1(0).ItemData(List1(0).ListIndex))
    Text2.Text = " "
    Text2.Text = rsTemp.Recordset.Fields("Ingredients")
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(0).Value = True Then
        btnTransfer.Enabled = True
        btnClear.Enabled = True
        List2.Enabled = True
    Else
        btnTransfer.Enabled = False
        btnClear.Enabled = False
        List2.Enabled = False
        Text1.SetFocus
    End If
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    List1(1).TopIndex = List1(0).TopIndex
End Sub
