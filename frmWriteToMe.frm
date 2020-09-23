VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmWriteToMe 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mail to Programme Developer"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.LaVolpeButton btnExit 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   5640
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1085
      BTNICON         =   "frmWriteToMe.frx":0000
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
      MICON           =   "frmWriteToMe.frx":015A
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
   Begin Project1.LaVolpeButton btnSend 
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   5640
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      BTNICON         =   "frmWriteToMe.frx":0176
      BTYPE           =   3
      TX              =   "&Send This Mail"
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
      MICON           =   "frmWriteToMe.frx":2928
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   4575
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   8070
         _Version        =   393217
         BackColor       =   16777152
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmWriteToMe.frx":2944
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "I would like to have the folowing errors corrected / new facilities added:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7695
      End
   End
End
Attribute VB_Name = "frmWriteToMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim boolOutlook As Boolean
Dim rsSupplier As Recordset
Dim rsLanguage As Recordset
Private Sub IsOutlookPresent()
Dim rVal As Variant, i As Integer
    rVal = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\mailto\shell\open\command", "")
    i = InStr(rVal, "OUTLOOK.EXE")
    If i <> 0 Then
        boolOutlook = True
    Else
        boolOutlook = False
    End If
End Sub

Private Sub ShowText()
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
                If IsNull(.Fields("btnSend")) Then
                    .Fields("btnSend") = btnSend.Caption
                Else
                    btnSend.Caption = .Fields("btnSend")
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
        .Fields("Label1") = Label1.Caption
        .Fields("btnSend") = btnSend.Caption
        .Fields("btnExit") = btnExit.Caption
        .Fields("Msg1") = "This mail is now send !"
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub


Private Sub btnSend_Click()
    On Error Resume Next
    IsOutlookPresent
    If boolOutlook Then
        If Len(RichTextBox1.Text) <> 0 Then
            Call SendOutlookMail("New message from program: Recipe", rsSupplier.Fields("SupplierEmailySysResponse"), RichTextBox1.Text)
        End If
    Else
        If IsWebConnected() Then
            If Not IsNull(rsSupplier.Fields("E-MailServerName")) Then
                On Error Resume Next
                Dim poSendMail As vbSendMail.clsSendMail
                Set poSendMail = New clsSendMail
                With poSendMail
                    .SMTPHost = Trim(rsSupplier.Fields("E-MailServerName"))
                    .from = Format(rsSupplier.Fields("EMail"))
                    .FromDisplayName = Format(rsSupplier.Fields("FirstName")) & " " & Format(rsSupplier.Fields("LastName"))
                    .Message = RichTextBox1.Text
                    .Recipient = Format(rsSupplier.Fields("SupplierEmailySysResponse"))
                    .Subject = "New message from program: Recipe"
                    .Send
                End With
            End If
            Set poSendMail = Nothing
            MsgBox rsLanguage.Fields("Msg1")
            Unload Me
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ShowText
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Set rsSupplier = m_dbData.OpenRecordset("MyRec")
    Set rsLanguage = m_dbData.OpenRecordset("frmWriteToMe")
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsSupplier.Close
    rsLanguage.Close
    Set frmWriteToMe = Nothing
End Sub
Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichTextBox1.SelText = vbTab
        KeyCode = 0
    End If
End Sub
