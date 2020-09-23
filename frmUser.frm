VERSION 5.00
Begin VB.Form frmUser 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Information"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6375
      Begin VB.ComboBox cboLanguagePrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LanguagePrint"
         DataSource      =   "rsMyRec"
         Height          =   315
         Left            =   3000
         TabIndex        =   25
         Top             =   3960
         Width           =   1815
      End
      Begin VB.ComboBox cboLanguageScreen 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LanguageScreen"
         DataSource      =   "rsMyRec"
         Height          =   315
         Left            =   3000
         TabIndex        =   23
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Data rsMyRec 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Recipes\Recipe.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "MyRec"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "E-MailAdress"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   10
         Left            =   3000
         TabIndex        =   21
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "EMail"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   9
         Left            =   3000
         TabIndex        =   19
         Top             =   3000
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Fax"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   8
         Left            =   3000
         TabIndex        =   17
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Telefon"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   7
         Left            =   3000
         TabIndex        =   15
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Country"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   6
         Left            =   3000
         TabIndex        =   13
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Town"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   5
         Left            =   3000
         TabIndex        =   11
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Zip"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   9
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Adress2"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   8
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Adress1"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   6
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LastName"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   4
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "FirstName"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Language on print:"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   26
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Language on screen:"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   24
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail Server Address:"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   22
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   20
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fax Number:"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Country:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Town:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Zip Code:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2655
      End
   End
   Begin Project1.LaVolpeButton btnExit 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      BTNICON         =   "frmUser.frx":0000
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
      MICON           =   "frmUser.frx":015A
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
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bookmarks1() As Variant, bookmarks2() As Variant
Dim rsCountry As Recordset
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
                For i = 0 To 11
                    If IsNull(.Fields(i + 2)) Then
                        .Fields(i + 2) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 2)
                    End If
                Next
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
        For i = 0 To 11
            .Fields(i + 2) = Label1(i).Caption
        Next
        .Fields("btnExit") = btnExit.Caption
        .Update
    End With
End Sub

Public Sub LoadCountry()
    On Error Resume Next
    cboLanguageScreen.Clear
    cboLanguagePrint.Clear
    With rsCountry
        .MoveLast
        .MoveFirst
        ReDim bookmarks1(.RecordCount)
        ReDim bookmarks2(.RecordCount)
        .Index = "PrimaryKey"
        Do While Not .EOF
            cboLanguageScreen.AddItem .Fields("Country")
            cboLanguageScreen.ItemData(cboLanguageScreen.NewIndex) = cboLanguageScreen.ListCount - 1
            bookmarks1(cboLanguageScreen.ListCount - 1) = .Bookmark
            cboLanguagePrint.AddItem .Fields("Country")
            cboLanguagePrint.ItemData(cboLanguagePrint.NewIndex) = cboLanguagePrint.ListCount - 1
            bookmarks2(cboLanguagePrint.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub


Private Sub cboLanguagePrint_Click()
    rsCountry.Bookmark = bookmarks2(cboLanguagePrint.ItemData(cboLanguagePrint.ListIndex))
    cboLanguagePrint.Text = rsCountry.Fields("CountryFix")
End Sub
Private Sub cboLanguageScreen_Click()
    On Error GoTo errChangeLanguage
    rsCountry.Bookmark = bookmarks1(cboLanguageScreen.ItemData(cboLanguageScreen.ListIndex))
    cboLanguageScreen.Text = rsCountry.Fields("CountryFix")
    
    'check if the language is present in the database
    If IsLanguagePresent(rsLanguage, Trim(rsCountry.Fields("CountryFix"))) Then
        m_FileExt = Trim(rsCountry.Fields("CountryFix"))
        frmRecipe.ReadText
        Exit Sub
    Else   'we did not have this language in the ProgramLang.mdb database
        MakeNewLanguage (Trim(rsCountry.Fields("CountryFix"))) 'make new recordset for each form in database
    End If
    m_FileExt = Trim(Trim(rsCountry.Fields("CountryFix")))
    ReadText    'show this form-text
    frmRecipe.ReadText
    Exit Sub
    
errChangeLanguage:
    Beep
    MsgBox Err.Description, vbExclamation, "Change Language"
    Err.Clear
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    LoadCountry
    rsMyRec.Refresh
    ReadText
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsMyRec.DatabaseName = m_sData
    Set rsCountry = m_dbData.OpenRecordset("Country")
    Set rsLanguage = m_dbData.OpenRecordset("frmUser")
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Form Load"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsMyRec.Recordset.Close
    rsCountry.Close
    rsLanguage.Close
    Erase bookmarks1
    Erase bookmarks2
    Set frmUser = Nothing
End Sub
