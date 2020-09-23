VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRecipe 
   BackColor       =   &H00404040&
   Caption         =   "Recipes"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14700
   Icon            =   "frmRecipe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   14700
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   35
      Top             =   8940
      Width           =   14700
      _ExtentX        =   25929
      _ExtentY        =   714
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17639
            MinWidth        =   17639
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:52"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "19.06.2002"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame6 
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   11640
      TabIndex        =   34
      Top             =   120
      Width           =   2895
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Notes"
         DataSource      =   "rsRecipe"
         Height          =   5655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search For:"
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   240
      TabIndex        =   32
      Top             =   360
      Width           =   3855
      Begin Project1.LaVolpeButton btnSearch 
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   450
         BTNICON         =   "frmRecipe.frx":08CA
         BTYPE           =   3
         TX              =   "Se&arch"
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
         MICON           =   "frmRecipe.frx":25D4
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
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3615
      End
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   11280
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Recipe Picture:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   11640
      TabIndex        =   31
      Top             =   6240
      Width           =   2895
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Picture"
         DataSource      =   "rsRecipe"
         Height          =   2175
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Recipe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   4440
      TabIndex        =   23
      Top             =   120
      Width           =   7095
      Begin VB.ComboBox cboRecipeCountry 
         BackColor       =   &H00FFFFC0&
         DataField       =   "RecipeCountry"
         DataSource      =   "rsRecipe"
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   0
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Author"
         DataSource      =   "rsRecipe"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1080
         Width           =   5415
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Instructions"
         DataSource      =   "rsRecipe"
         Height          =   4335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   4200
         Width           =   6855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Ingredients"
         DataSource      =   "rsRecipe"
         Height          =   1815
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1560
         Width           =   5415
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Name"
         DataSource      =   "rsRecipe"
         Height          =   345
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   2
         Top             =   720
         Width           =   5415
      End
      Begin VB.Data rsRecipe 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Recipes\Recipe.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Recipe"
         Top             =   120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Timer Timer1 
         Interval        =   60
         Left            =   5160
         Top             =   120
      End
      Begin VB.ComboBox cboRecipeType 
         BackColor       =   &H00FFFFC0&
         DataField       =   "RecipeType"
         DataSource      =   "rsRecipe"
         Height          =   315
         Index           =   1
         Left            =   4920
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Serves"
         DataSource      =   "rsRecipe"
         Height          =   345
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   5
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(If zero (0) the serves = unspecified)"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   11
         Left            =   2760
         TabIndex        =   36
         Top             =   3600
         Width           =   4185
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recipe Name:"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingredients:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   1305
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preparation Instructions:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   27
         Top             =   3960
         Width           =   3375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   8
         Left            =   3915
         TabIndex        =   26
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serves:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   24
         Top             =   3600
         Width           =   1305
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Database:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4095
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   4905
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   4905
         Index           =   1
         Left            =   1200
         TabIndex        =   19
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   3855
         Begin VB.ComboBox cboRecipeCountry 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   2175
         End
         Begin VB.ComboBox cboRecipeType 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Index           =   0
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   960
            Width           =   2175
         End
         Begin Project1.LaVolpeButton btnSearchType 
            Height          =   255
            Left            =   2280
            TabIndex        =   15
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            BTNICON         =   "frmRecipe.frx":25F0
            BTYPE           =   3
            TX              =   "&Search"
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
            MICON           =   "frmRecipe.frx":42FA
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
         Begin Project1.LaVolpeButton btnSearchCountry 
            Height          =   255
            Left            =   2280
            TabIndex        =   16
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            BTNICON         =   "frmRecipe.frx":4316
            BTYPE           =   3
            TX              =   "S&earch"
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
            MICON           =   "frmRecipe.frx":6020
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recipe Type:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recipe Country:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   1140
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
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
         Height          =   195
         Index           =   6
         Left            =   300
         TabIndex        =   22
         Top             =   2640
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recipe"
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
         Height          =   195
         Index           =   7
         Left            =   2265
         TabIndex        =   21
         Top             =   2640
         Width           =   645
      End
   End
   Begin Project1.LaVolpeButton btnDelete 
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   8280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTNICON         =   "frmRecipe.frx":603C
      BTYPE           =   3
      TX              =   "&Delete Recipe"
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
      MICON           =   "frmRecipe.frx":6196
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
      Left            =   120
      TabIndex        =   12
      Top             =   8280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BTNICON         =   "frmRecipe.frx":61B2
      BTYPE           =   3
      TX              =   "&New Recipe"
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
      MICON           =   "frmRecipe.frx":6884
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
   Begin VB.Menu mnuFiles 
      Caption         =   "&Files"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuRecipeType 
         Caption         =   "Recipe Type"
      End
      Begin VB.Menu mnuIngredients 
         Caption         =   "Ingredients"
      End
      Begin VB.Menu mnuDimension 
         Caption         =   "Dimensions"
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLanguage 
         Caption         =   "Language"
         Begin VB.Menu mnuNewLanguage 
            Caption         =   "Select New Language"
         End
         Begin VB.Menu mnuScreenText 
            Caption         =   "Change Screen text"
         End
      End
      Begin VB.Menu space3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUser 
         Caption         =   "User Information"
      End
   End
   Begin VB.Menu mnuMail 
      Caption         =   "&Mail"
      Begin VB.Menu mnuMailSend 
         Caption         =   "Send this Recipe"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
      Begin VB.Menu mnuPrintSetUp 
         Caption         =   "Printer Set-Up"
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintRecipe 
         Caption         =   "Print this Recipe"
      End
   End
   Begin VB.Menu mnuPicture 
      Caption         =   "&Picture"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy Picture"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste Picture"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut Picture"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Load Picture from disk"
      End
      Begin VB.Menu mnuPrintPicture 
         Caption         =   "Print Picture"
      End
   End
   Begin VB.Menu mnuInternet 
      Caption         =   "&Internet"
      Begin VB.Menu mnuLinks 
         Caption         =   "Recipe Links"
      End
   End
   Begin VB.Menu mnuRecipe 
      Caption         =   "&Recipe"
      Begin VB.Menu mnuSearchRecipe 
         Caption         =   "Search for Recipe"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAboutMe 
         Caption         =   "About"
      End
      Begin VB.Menu mnuWriteToMe 
         Caption         =   "Write to Me"
      End
   End
End
Attribute VB_Name = "frmRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bookmarks1() As Variant, bNewRecord As Boolean, i As Integer
Dim dbTemp As Database
Dim rsRecipeType As Recordset
Dim rsMyRec As Recordset
Dim rsLanguage As Recordset
Dim rsCountry As Recordset
Private Sub RecipePrint()
    On Error Resume Next
    m_LeftMargin = 20
    m_TopMargin = 20
    m_BottomMargin = 20
    m_RightMargin = 10
    
    Printer.ScaleMode = vbMillimeters
    Printer.CurrentY = m_TopMargin
    Printer.FontSize = 12
    Printer.CurrentX = m_LeftMargin
    Printer.Print Label1(0).Caption & " " & Trim(Text1.Text) & " (" & cboRecipeCountry(1).Text & ")"    'recipe name
    Printer.CurrentX = m_LeftMargin
    Printer.Print Label1(3).Caption & " " & Trim(Text2.Text)    'author
    Printer.CurrentX = m_LeftMargin
    Printer.Print Label1(10).Caption & " " & Trim(Text5.Text)   'serves
    Printer.Print
    Printer.CurrentX = m_LeftMargin
    Printer.Print Label1(4).Caption & " "
    Printer.CurrentX = m_LeftMargin
    Call PrintLongText(Text3.Text, m_LeftMargin)
    Printer.Print
    Printer.Print
    Printer.CurrentX = m_LeftMargin
    Printer.Print Label1(5).Caption & " "
    Printer.CurrentX = m_LeftMargin
    Call PrintLongText(Text4.Text, m_LeftMargin)
    Printer.Print
    Printer.Print
    Printer.CurrentX = m_LeftMargin
    Printer.Print Frame6.Caption & " "
    Printer.CurrentX = m_LeftMargin
    Call PrintLongText(Text7.Text, m_LeftMargin)
    Printer.EndDoc
End Sub


Private Function SearchRecipe() As Boolean
Dim Sql As String
    On Error GoTo errSearchRecipe
    Sql = "SELECT * FROM Recipe WHERE InStr(Ingredients, "
    Sql = Sql & Chr(34) & CStr(Text6.Text) & Chr(34)
    Sql = Sql & ") > 0 ORDER BY Name"
    
    rsRecipe.RecordSource = Sql
    rsRecipe.Refresh
    SearchRecipe = True
    Exit Function
    
errSearchRecipe:
    Err.Clear
    SearchRecipe = False
End Function


Public Sub LoadCountry()
    On Error Resume Next
    cboRecipeCountry(0).Clear
    cboRecipeCountry(1).Clear
    With rsCountry
        .MoveFirst
        .Index = "PrimaryKey"
        Do While Not .EOF
            cboRecipeCountry(0).AddItem .Fields("Country")
            cboRecipeCountry(1).AddItem .Fields("Country")
        .MoveNext
        Loop
    End With
End Sub

Private Sub LoadList1()
    On Error Resume Next
    List1(0).Clear
    List1(1).Clear
    If rsRecipe.Recordset.RecordCount = 0 Then Exit Sub
    With rsRecipe.Recordset
        If Not .BOF And Not .EOF Then
            .MoveLast
            ReDim bookmarks1(.RecordCount)
            .MoveFirst
            Do While Not .EOF
                List1(0).AddItem .Fields("RecipeCountry")
                List1(0).ItemData(List1(0).NewIndex) = List1(0).ListCount - 1
                bookmarks1(List1(0).ListCount - 1) = .Bookmark
                List1(1).AddItem .Fields("Name")
            .MoveNext
            Loop
        End If
    End With
End Sub
Public Sub LoadRecipeType()
    On Error Resume Next
    cboRecipeType(0).Clear
    cboRecipeType(1).Clear
    With rsRecipeType
        .MoveFirst
        Do While Not .EOF
            cboRecipeType(0).AddItem .Fields("TypeName")
            cboRecipeType(1).AddItem .Fields("TypeName")
        .MoveNext
        Loop
    End With
End Sub

Private Sub btnNew_Click()
    On Error Resume Next
    bNewRecord = True
    rsRecipe.Recordset.AddNew
    cboRecipeCountry(1).SetFocus
End Sub

Public Sub ReadText()
    'find YOUR sLanguage text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then
                .Edit
                For i = 0 To 11
                    If IsNull(.Fields(i + 2)) Then
                        .Fields(i + 2) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 2)
                    End If
                Next
                If IsNull(.Fields("Form")) Then
                    .Fields("Form") = Me.Caption
                Else
                    Me.Caption = .Fields("Form")
                    Me.Caption = Me.Caption & "  Version: " & App.Major & "." & App.Minor & "." & App.Revision
                End If
                If IsNull(.Fields("btnSearchCountry")) Then
                    .Fields("btnSearchCountry") = btnSearchCountry.Caption
                Else
                    btnSearchCountry.Caption = .Fields("btnSearchCountry")
                End If
                If IsNull(.Fields("btnSearchType")) Then
                    .Fields("btnSearchType") = btnSearchType.Caption
                Else
                    btnSearchType.Caption = .Fields("btnSearchType")
                End If
                If IsNull(.Fields("btnSearch")) Then
                    .Fields("btnSearch") = btnSearch.Caption
                Else
                    btnSearch.Caption = .Fields("btnSearch")
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
                If IsNull(.Fields("mnuPicture")) Then
                    .Fields("mnuPicture") = mnuPicture.Caption
                Else
                    mnuPicture.Caption = .Fields("mnuPicture")
                End If
                If IsNull(.Fields("mnuPaste")) Then
                    .Fields("mnuPaste") = mnuPaste.Caption
                Else
                    mnuPaste.Caption = .Fields("mnuPaste")
                End If
                If IsNull(.Fields("mnuCopy")) Then
                    .Fields("mnuCopy") = mnuCopy.Caption
                Else
                    mnuCopy.Caption = .Fields("mnuCopy")
                End If
                If IsNull(.Fields("mnuCut")) Then
                    .Fields("mnuCut") = mnuCut.Caption
                Else
                    mnuCut.Caption = .Fields("mnuCut")
                End If
                If IsNull(.Fields("mnuPrintPicture")) Then
                    .Fields("mnuPrintPicture") = mnuPrintPicture.Caption
                Else
                    mnuPrintPicture.Caption = .Fields("mnuPrintPicture")
                End If
                If IsNull(.Fields("mnuOpen")) Then
                    .Fields("mnuOpen") = mnuOpen.Caption
                Else
                    mnuOpen.Caption = .Fields("mnuOpen")
                End If
                If IsNull(.Fields("mnuFiles")) Then
                    .Fields("mnuFiles") = mnuFiles.Caption
                Else
                    mnuFiles.Caption = .Fields("mnuFiles")
                End If
                If IsNull(.Fields("mnuExit")) Then
                    .Fields("mnuExit") = mnuExit.Caption
                Else
                    mnuExit.Caption = .Fields("mnuExit")
                End If
                If IsNull(.Fields("mnuDatabase")) Then
                    .Fields("mnuDatabase") = mnuDatabase.Caption
                Else
                    mnuDatabase.Caption = .Fields("mnuDatabase")
                End If
                If IsNull(.Fields("mnuRecipeType")) Then
                    .Fields("mnuRecipeType") = mnuRecipeType.Caption
                Else
                    mnuRecipeType.Caption = .Fields("mnuRecipeType")
                End If
                If IsNull(.Fields("mnuIngredients")) Then
                    .Fields("mnuIngredients") = mnuIngredients.Caption
                Else
                    mnuIngredients.Caption = .Fields("mnuIngredients")
                End If
                If IsNull(.Fields("mnuLanguage")) Then
                    .Fields("mnuLanguage") = mnuLanguage.Caption
                Else
                    mnuLanguage.Caption = .Fields("mnuLanguage")
                End If
                If IsNull(.Fields("mnuNewLanguage")) Then
                    .Fields("mnuNewLanguage") = mnuNewLanguage.Caption
                Else
                    mnuNewLanguage.Caption = .Fields("mnuNewLanguage")
                End If
                If IsNull(.Fields("mnuScreenText")) Then
                    .Fields("mnuScreenText") = mnuScreenText.Caption
                Else
                    mnuScreenText.Caption = .Fields("mnuScreenText")
                End If
                If IsNull(.Fields("mnuDimension")) Then
                    .Fields("mnuDimension") = mnuDimension.Caption
                Else
                    mnuDimension.Caption = .Fields("mnuDimension")
                End If
                If IsNull(.Fields("mnuMail")) Then
                    .Fields("mnuMail") = mnuMail.Caption
                Else
                    mnuMail.Caption = .Fields("mnuMail")
                End If
                If IsNull(.Fields("mnuMailSend")) Then
                    .Fields("mnuMailSend") = mnuMailSend.Caption
                Else
                    mnuMailSend.Caption = .Fields("mnuMailSend")
                End If
                If IsNull(.Fields("mnuPrint")) Then
                    .Fields("mnuPrint") = mnuPrint.Caption
                Else
                    mnuPrint.Caption = .Fields("mnuPrint")
                End If
                If IsNull(.Fields("mnuPrintSetUp")) Then
                    .Fields("mnuPrintSetUp") = mnuPrintSetUp.Caption
                Else
                    mnuPrintSetUp.Caption = .Fields("mnuPrintSetUp")
                End If
                If IsNull(.Fields("mnuPrintRecipe")) Then
                    .Fields("mnuPrintRecipe") = mnuPrintRecipe.Caption
                Else
                    mnuPrintRecipe.Caption = .Fields("mnuPrintRecipe")
                End If
                If IsNull(.Fields("mnuUser")) Then
                    .Fields("mnuUser") = mnuUser.Caption
                Else
                    mnuUser.Caption = .Fields("mnuUser")
                End If
                If IsNull(.Fields("mnuAbout")) Then
                    .Fields("mnuAbout") = mnuAbout.Caption
                Else
                    mnuAbout.Caption = .Fields("mnuAbout")
                End If
                If IsNull(.Fields("mnuAboutMe")) Then
                    .Fields("mnuAboutMe") = mnuAboutMe.Caption
                Else
                    mnuAboutMe.Caption = .Fields("mnuAboutMe")
                End If
                If IsNull(.Fields("mnuWriteToMe")) Then
                    .Fields("mnuWriteToMe") = mnuWriteToMe.Caption
                Else
                    mnuWriteToMe.Caption = .Fields("mnuWriteToMe")
                End If
                If IsNull(.Fields("mnuInternet")) Then
                    .Fields("mnuInternet") = mnuInternet.Caption
                Else
                    mnuInternet.Caption = .Fields("mnuInternet")
                End If
                If IsNull(.Fields("mnuLinks")) Then
                    .Fields("mnuLinks") = mnuLinks.Caption
                Else
                    mnuLinks.Caption = .Fields("mnuLinks")
                End If
                If IsNull(.Fields("mnuRecipe")) Then
                    .Fields("mnuRecipe") = mnuRecipe.Caption
                Else
                    mnuRecipe.Caption = .Fields("mnuRecipe")
                End If
                If IsNull(.Fields("mnuSearchRecipe")) Then
                    .Fields("mnuSearchRecipe") = mnuSearchRecipe.Caption
                Else
                    mnuSearchRecipe.Caption = .Fields("mnuSearchRecipe")
                End If
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
                If IsNull(.Fields("Frame5")) Then
                    .Fields("Frame5") = Frame5.Caption
                Else
                    Frame5.Caption = .Fields("Frame5")
                End If
                If IsNull(.Fields("Frame6")) Then
                    .Fields("Frame6") = Frame6.Caption
                Else
                    Frame6.Caption = .Fields("Frame6")
                End If
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        For i = 0 To 11
            .Fields(i + 1) = Label1(i).Caption
        Next
        .Fields("btnSearchCountry") = btnSearchCountry.Caption
        .Fields("btnSearchType") = btnSearchType.Caption
        .Fields("btnSearch") = btnSearch.Caption
        .Fields("btnNew") = btnNew.Caption
        .Fields("btnDelete") = btnDelete.Caption
        .Fields("mnuPicture") = mnuPicture.Caption
        .Fields("mnuPaste") = mnuPaste.Caption
        .Fields("mnuCopy") = mnuCopy.Caption
        .Fields("mnuCut") = mnuCut.Caption
        .Fields("mnuPrintPicture") = mnuPrintPicture.Caption
        .Fields("mnuOpen") = mnuOpen.Caption
        .Fields("mnuFiles") = mnuFiles.Caption
        .Fields("mnuExit") = mnuExit.Caption
        .Fields("mnuDatabase") = mnuDatabase.Caption
        .Fields("mnuRecipeType") = mnuRecipeType.Caption
        .Fields("mnuIngredients") = mnuIngredients.Caption
        .Fields("mnuLanguage") = mnuLanguage.Caption
        .Fields("mnuNewLanguage") = mnuNewLanguage.Caption
        .Fields("mnuScreenText") = mnuScreenText.Caption
        .Fields("mnuDimension") = mnuDimension.Caption
        .Fields("mnuMail") = mnuMail.Caption
        .Fields("mnuMailSend") = mnuMailSend.Caption
        .Fields("mnuPrint") = mnuPrint.Caption
        .Fields("mnuPrintSetUp") = mnuPrintSetUp.Caption
        .Fields("mnuPrintRecipe") = mnuPrintRecipe.Caption
        .Fields("mnuUser") = mnuUser.Caption
        .Fields("mnuAbout") = mnuAbout.Caption
        .Fields("mnuAboutMe") = mnuAboutMe.Caption
        .Fields("mnuWriteToMe") = mnuWriteToMe.Caption
        .Fields("mnuInternet") = mnuInternet.Caption
        .Fields("mnuLinks") = mnuLinks.Caption
        .Fields("mnuRecipe") = mnuRecipe.Caption
        .Fields("mnuSearchRecipe") = mnuSearchRecipe.Caption
        .Fields("Frame1") = Frame1.Caption
        .Fields("Frame2") = Frame2.Caption
        .Fields("Frame3") = Frame2.Caption
        .Fields("Frame5") = Frame5.Caption
        .Fields("Frame6") = Frame6.Caption
    End With
End Sub
Private Function SelectRecipe() As Boolean
Dim Sql As String
    On Error GoTo errSelectRecipe
    Sql = "SELECT * FROM Recipe WHERE RecipeType ="
    Sql = Sql & Chr(34) & cboRecipeType(0).Text & Chr(34)
    Sql = Sql & " ORDER BY Name"
    rsRecipe.RecordSource = Sql
    rsRecipe.Refresh
    SelectRecipe = True
    Exit Function
    
errSelectRecipe:
    Err.Clear
    SelectRecipe = False
End Function
Private Function SelectRecipeCountry() As Boolean
Dim Sql As String
    On Error GoTo errSelectRecipeCountry
    Sql = "SELECT * FROM Recipe WHERE Trim(RecipeCountry) ="
    Sql = Sql & Chr(34) & Trim(cboRecipeCountry(0).Text) & Chr(34)
    Sql = Sql & " ORDER BY Name"
    rsRecipe.RecordSource = Sql
    rsRecipe.Refresh
    SelectRecipeCountry = True
    Exit Function
    
errSelectRecipeCountry:
    Err.Clear
    SelectRecipeCountry = False
End Function

Public Sub ScanPicture()
Dim Ret As Long, T As Variant
    On Error Resume Next
    Ret = TWAIN_AcquireToClipboard(Me.hWnd, T)
    Image1.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub btnSearch_Click()
    List1(0).Clear
    List1(1).Clear
    If SearchRecipe Then
        LoadList1
    End If
End Sub
Private Sub btnSearchCountry_Click()
    On Error Resume Next
    List1(0).Clear
    List1(1).Clear
    If SelectRecipeCountry Then
        LoadList1
        List1(0).ListIndex = 0
    End If
End Sub

Private Sub btnSearchType_Click()
    On Error Resume Next
    List1(0).Clear
    List1(1).Clear
    If SelectRecipe Then
        LoadList1
        List1(0).ListIndex = 0
    End If
End Sub

Private Sub ActivateMe()
    On Error Resume Next
    rsRecipe.Refresh
    m_FileExt = rsMyRec.Fields("LanguageScreen")
    LoadRecipeType
    LoadCountry
    ReadText
    If SelectRecipeCountry Then
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Me.WindowState = vbMaximized
    StatusBar1.Panels(1).Text = App.Path
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    m_sData = App.Path & "\Recipe.mdb"
    rsRecipe.DatabaseName = m_sData
    Set m_dbData = OpenDatabase(m_sData)
    Set rsRecipeType = m_dbData.OpenRecordset("RecipeType")
    Set rsCountry = m_dbData.OpenRecordset("Country")
    Set rsMyRec = m_dbData.OpenRecordset("MyRec")
    Set rsLanguage = m_dbData.OpenRecordset("frmRecipe")
    ActivateMe
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "LoadForm"
    Err.Clear
    Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    rsRecipe.Recordset.Close
    rsRecipeType.Close
    rsCountry.Close
    rsLanguage.Close
    m_dbData.Close
    Erase bookmarks1
End Sub
Private Sub Form_Resize()
    ResizeForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmRecipe = Nothing
End Sub



Private Sub List1_Click(Index As Integer)
    On Error Resume Next
    i = List1(Index).ListIndex
    List1(0).ListIndex = i
    List1(1).ListIndex = i
    With rsRecipe.Recordset
        .Bookmark = bookmarks1(List1(0).ItemData(List1(0).ListIndex))
        cboRecipeCountry(0).Text = .Fields("RecipeCountry")
        cboRecipeType(0).Text = .Fields("RecipeType")
    End With
End Sub
Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim WordHeight As Long
    'the height of the default text (you will have to change this if you change the font size)
    WordHeight = 195
    'go through the loop until you get to the file
    For i = 0 To List1(1).ListCount - 1
        'check to what line the text is over (you need to go through the whole list in case you've
        'scrolled down
        If Y > WordHeight * (i - List1(1).TopIndex) _
            And Y < (WordHeight * i + WordHeight) Then
            'set the tooltiptext to the list box value
            List1(1).ToolTipText = List1(1).List(i)
            'see if your in "empty space"
        ElseIf Y > (WordHeight * i + WordHeight) Then
            List1(1).ToolTipText = "Empty space"
        End If
    Next i
End Sub

Private Sub mnuAboutMe_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuCopy_Click()
    On Error GoTo errCopy
    Clipboard.Clear
    Clipboard.SetData Image1.Picture, vbCFDIB
    Exit Sub
    
errCopy:
    Beep
    MsgBox Err.Description, vbCritical, "Copy Picture to Clipboard"
    Err.Clear
End Sub

Private Sub mnuCut_Click()
    On Error GoTo errCut
    Clipboard.Clear
    Clipboard.SetData Image1.Picture, vbCFDIB
    Image1.Picture = LoadPicture()
    Exit Sub
    
errCut:
    Beep
    MsgBox Err.Description, vbExclamation, "Cut picture"
    Err.Clear
End Sub

Private Sub mnuDimension_Click()
    frmDimension.Show 1
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuIngredients_Click()
    frmIngredients.Show 1
End Sub


Private Sub mnuLinks_Click()
    frmLinks.Show 1
End Sub

Private Sub mnuMailSend_Click()
    With frmMail
        .LoadMessage
        .Show 1
    End With
End Sub

Private Sub mnuNewLanguage_Click()
    frmCountry.Show 1
End Sub

Private Sub mnuOpen_Click()
    'Read the picture from disk
    With Cmd1
        .FileName = ""
        .DialogTitle = "Load Picture from disk"
        .Filter = "Pictures (*.bmp; *.pcx;*.jpg)|*.bmp;*.pcx;*.jpg"
        .FilterIndex = 1
        .Action = 1
    End With
    
    On Error GoTo errReadFromFile
    Set Image1.Picture = LoadPicture(Cmd1.FileName)
    Exit Sub
    
errReadFromFile:
    Me.MousePointer = Default
    Beep
    MsgBox Err.Description, vbCritical, " Read File"
    Err.Clear
End Sub

Private Sub mnuPaste_Click()
    On Error GoTo errPaste
    Image1.Picture = Clipboard.GetData(vbCFDIB)
    Exit Sub
    
errPaste:
    Beep
    MsgBox Err.Description, vbCritical, "Error Clipboard"
    Err.Clear
End Sub

Private Sub mnuPrintPicture_Click()
    If Not IsNull(rsRecipe.Recordset.Fields("Picture")) Then
        PrintPictureToFitPage Printer, Image1.Picture
    End If
End Sub

Private Sub mnuPrintRecipe_Click()
    RecipePrint
End Sub

Private Sub mnuPrintSetUp_Click()
    Cmd1.ShowPrinter
End Sub

Private Sub mnuRecipeType_Click()
    frmRecipeTypes.Show 1
End Sub

Private Sub mnuScreenText_Click()
    frmScreenLanguage.Show 1
End Sub

Private Sub mnuSearchRecipe_Click()
    On Error Resume Next
    With frmHaveIngredients
        .cboRecipeType.Clear
            rsRecipeType.MoveFirst
            Do While Not rsRecipeType.EOF
                .cboRecipeType.AddItem rsRecipeType.Fields("TypeName")
            rsRecipeType.MoveNext
            Loop
        .Show 1
    End With
End Sub

Private Sub mnuUser_Click()
    frmUser.Show 1
End Sub

Private Sub mnuWriteToMe_Click()
    frmWriteToMe.Show 1
End Sub

Private Sub Text1_GotFocus()
    onGotFocus
End Sub

Private Sub Text1_LostFocus()
    If bNewRecord Then
        With rsRecipe.Recordset
            .Fields("RecipeCountry") = Trim(cboRecipeCountry(1).Text)
            .Fields("RecipeType") = Trim(cboRecipeType(1).Text)
            .Fields("Name") = Trim(Text1.Text)
            .Update
            LoadList1
            .Bookmark = .LastModified
        End With
        bNewRecord = False
    End If
End Sub

Private Sub Text2_GotFocus()
    onGotFocus
End Sub


Private Sub Text5_GotFocus()
    onGotFocus
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    List1(1).TopIndex = List1(0).TopIndex
End Sub


