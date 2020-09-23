VERSION 5.00
Begin VB.Form frmMail 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mail "
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.LaVolpeButton btnExit 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTNICON         =   "frmMail.frx":0000
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
      MICON           =   "frmMail.frx":015A
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
   Begin Project1.LaVolpeButton btnSendMail 
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   6840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      BTNICON         =   "frmMail.frx":0176
      BTYPE           =   3
      TX              =   "&Send this mail"
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
      MICON           =   "frmMail.frx":2928
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
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   8055
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   4485
         Index           =   3
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Mail Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.ComboBox cboAdr 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   2760
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   4335
      End
      Begin Project1.LaVolpeButton btnMailTo 
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "To... (Outlook address)"
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
         MICON           =   "frmMail.frx":2944
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
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   2
         Left            =   2760
         TabIndex        =   4
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   1
         Left            =   2760
         TabIndex        =   2
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mail from:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ol As Object
Dim olns As Object
Dim objFolder As Object
Dim objAllContacts As Object
Dim Contact As Object
Dim boolOutlook As Boolean
Dim rsMyRec As Recordset
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
                If IsNull(.Fields("label1(0)")) Then
                    .Fields("label1(0)") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("label1(0)")
                End If
                If IsNull(.Fields("label1(1)")) Then
                    .Fields("label1(1)") = Label1(1).Caption
                Else
                    Label1(1).Caption = .Fields("label1(1)")
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
                If IsNull(.Fields("btnMailTo")) Then
                    .Fields("btnMailTo") = btnMailTo.Caption
                Else
                    btnMailTo.Caption = .Fields("btnMailTo")
                End If
                If IsNull(.Fields("btnSendMail")) Then
                    .Fields("btnSendMail") = btnSendMail.Caption
                Else
                    btnSendMail.Caption = .Fields("btnSendMail")
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
        .Fields("label1(0)") = Label1(0).Caption
        .Fields("label1(1)") = Label1(1).Caption
        .Fields("Frame1") = Frame1.Caption
        .Fields("Frame2") = Frame2.Caption
        .Fields("btnMailTo") = btnMailTo.Caption
        .Fields("btnSendMail") = btnSendMail.Caption
        .Fields("btnExit") = btnExit.Caption
        .Fields("Msg1") = "This Mail is now send !"
        .Update
    End With
End Sub

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
Private Sub LoadContacts()
    On Error GoTo errLoadContact
    If boolOutlook Then
        ' Set the application object
        Set ol = New Outlook.Application
        ' Set the namespace object
        Set olns = ol.GetNamespace("MAPI")
        ' Set the default Contacts folder
        Set objFolder = olns.GetDefaultFolder(olFolderContacts)
        ' Set objAllContacts = the collection of all contacts
        Set objAllContacts = objFolder.Items
        
        cboAdr.Clear
        ' Loop through each contact
        For Each Contact In objAllContacts
           'Display the Fullname field for the contact
           cboAdr.AddItem Contact.FullName
        Next
    End If
    Exit Sub
    
errLoadContact:
    Err.Clear
End Sub


Public Sub LoadMessage()
    On Error Resume Next
    With frmRecipe
        Text1(3).Text = .Label1(9).Caption & "  " & .cboRecipeCountry(1).Text   'country
        Text1(3).Text = Text1(3).Text & vbCrLf
        Text1(3).Text = Text1(3).Text & .Label1(0).Caption & "  " & .Text1.Text   'recipe name
        Text1(3).Text = Text1(3).Text & vbCrLf
        Text1(3).Text = Text1(3).Text & .Label1(3).Caption & "  " & .Text2.Text   'author
        Text1(3).Text = Text1(3).Text & vbCrLf & vbCrLf
        Text1(3).Text = Text1(3).Text & .Label1(4).Caption & " " & .Text3.Text  'ingredients
        Text1(3).Text = Text1(3).Text & vbCrLf & vbCrLf
        Text1(3).Text = Text1(3).Text & .Label1(10).Caption & "  " & .Text5.Text   'serves
        Text1(3).Text = Text1(3).Text & vbCrLf & vbCrLf
        Text1(3).Text = Text1(3).Text & .Label1(5).Caption & " " & .Text4.Text    'preperations
        Text1(3).Text = Text1(3).Text & vbCrLf & vbCrLf
        Text1(3).Text = Text1(3).Text & .Frame6.Caption & " " & .Text7.Text    'notes
        Text1(1).Text = .Caption
    End With
End Sub
Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnMailTo_Click()
    If boolOutlook Then cboAdr.Visible = True
End Sub

Private Sub btnSendMail_Click()
    If boolOutlook Then
        Dim clsOutlook As cOutlookSendMail
        On Error Resume Next
        Set clsOutlook = New cOutlookSendMail
        With clsOutlook
            .StartOutlook
            .CreateNewMail
            .Recipient_TO = Text1(0).Text
            .Subject = Text1(1).Text
            .Body = Text1(3).Text
            .SendMail
            .CloseOutlook
        End With
        Set clsOutlook = Nothing
        MsgBox rsLanguage.Fields("Msg1")
        Unload Me
    Else
        If IsWebConnected() Then
            If Not IsNull(rsMyRec.Fields("E-MailServerName")) Then
                On Error Resume Next
                Dim poSendMail As vbSendMail.clsSendMail
                Set poSendMail = New clsSendMail
                With poSendMail
                    .SMTPHost = Trim(rsMyRec.Fields("E-MailServerName"))
                    .from = Text1(2).Text
                    .FromDisplayName = Text1(3).Text
                    .Message = Text1(3).Text
                    .Recipient = Text1(0).Text
                    .Subject = Text1(1).Text
                    .Send
                End With
            End If
            Set poSendMail = Nothing
            MsgBox rsLanguage.Fields("Msg1")
            Unload Me
        End If
    End If
End Sub

Private Sub cboAdr_Click()
Dim sFilter As String
    sFilter = "[FullName] = """ & cboAdr.List(cboAdr.ListIndex) & """"
    Set Contact = objAllContacts.Find(sFilter)
    If Contact Is Nothing Then ' the Find failed
       MsgBox "Not Found"
    Else
        If Contact.Email1Address <> "" Then
            Text1(0).Text = Contact.Email1Address
        End If
    End If
    cboAdr.Visible = False
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    IsOutlookPresent
    If boolOutlook Then
        LoadContacts
    End If
    If Not IsNull(rsMyRec.Fields("EMail")) Then
        Text1(2).Text = rsMyRec.Fields("EMail")
    End If
    ReadText
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Set rsMyRec = m_dbData.OpenRecordset("MyRec")
    Set rsLanguage = m_dbData.OpenRecordset("frmMail")
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "LoadForm"
    Err.Clear
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set ol = Nothing
    Set olns = Nothing
    Set objFolder = Nothing
    Set objAllContacts = Nothing
    Set Contact = Nothing
    rsMyRec.Close
    rsLanguage.Close
    Set frmMail = Nothing
End Sub
