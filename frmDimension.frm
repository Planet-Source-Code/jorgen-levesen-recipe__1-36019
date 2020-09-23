VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDimension 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dimensions"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data rsDimension 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\Recipes\Recipe.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Dimension"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin Project1.LaVolpeButton btnExit 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      BTNICON         =   "frmDimension.frx":0000
      BTYPE           =   3
      TX              =   "&Exit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      MICON           =   "frmDimension.frx":015A
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
   Begin MSDBGrid.DBGrid Grid1 
      Bindings        =   "frmDimension.frx":0176
      Height          =   6135
      Left            =   120
      OleObjectBlob   =   "frmDimension.frx":0190
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmDimension"
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
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsDimension.Refresh
    ReadText
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsDimension.DatabaseName = m_sData
    Set rsLanguage = m_dbData.OpenRecordset("frmDimension")
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "LoadForm"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsDimension.Recordset.Close
    rsLanguage.Close
    Set frmDimension = Nothing
End Sub
