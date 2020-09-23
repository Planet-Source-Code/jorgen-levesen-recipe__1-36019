VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRecipeTypes 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recipe Types"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.LaVolpeButton btnExit 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   6120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      BTNICON         =   "frmRecipeTypes.frx":0000
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
      MICON           =   "frmRecipeTypes.frx":015A
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
   Begin VB.Data rsRecipeTypes 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\MasterHom\RECIPE.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "RecipeType"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid Grid1 
      Bindings        =   "frmRecipeTypes.frx":0176
      Height          =   5775
      Left            =   240
      OleObjectBlob   =   "frmRecipeTypes.frx":0192
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmRecipeTypes"
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
        .Fields("btnExit") = btnExit.Caption
        .Update
    End With
End Sub


Private Sub btnExit_Click()
    frmRecipe.LoadRecipeType
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsRecipeTypes.Refresh
    ReadText
End Sub
Private Sub Form_Load()
    On Error Resume Next
    rsRecipeTypes.DatabaseName = m_sData
    Set rsLanguage = m_dbData.OpenRecordset("frmRecipeTypes")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsRecipeTypes.Recordset.Close
    rsLanguage.Close
    Set frmRecipeTypes = Nothing
End Sub
