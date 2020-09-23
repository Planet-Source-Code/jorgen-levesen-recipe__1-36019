Attribute VB_Name = "basRecipe"
Option Explicit
Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, _
    ByVal dwReserved As Long) As Long

Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, lpcbData As Long) As Long
    
Private Const CONNECT_LAN As Long = &H2
Private Const CONNECT_MODEM As Long = &H1
Private Const CONNECT_PROXY As Long = &H4
Private Const CONNECT_OFFLINE As Long = &H20
Private Const CONNECT_CONFIGURED As Long = &H40

Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1 ' Unicode nul terminated String
Public Const REG_DWORD = 4 ' 32-bit number

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Declare Function GetModuleHandle Lib _
    "Kernel" (ByVal lpModuleName As String) As Integer
    
Public Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp&, ByVal wPixTypes&) As Long
Public Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp&) As Long
Public Declare Function GetWindowsDirectory& Lib "kernel32" Alias "GetWindowsDirectoryA" _
   (ByVal lpBuffer As String, ByVal nSize As Long)
   
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal _
    bRevert As Long) As Long
    
Public Declare Function ShellExecute Lib _
    "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Global wdApp As Word.Application
Global m_sData As String
Global m_dbData As Database
Global m_iWhichForm As Integer
Global m_FileExt As String
Global m_FilePrt As String
Global m_strMemo As String
Global i As Integer

Global m_LeftMargin As Integer
Global m_TopMargin As Integer
Global m_RightMargin As Integer
Global m_BottomMargin As Integer

Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const MB_YESNO = 4              ' Yes and No buttons
Global Const MB_RETRYCANCEL = 5        ' Retry and Cancel buttons

Global Const MB_ICONSTOP = 16          ' Critical message
Global Const MB_ICONQUESTION = 32      ' Warning query
Global Const MB_ICONEXCLAMATION = 48   ' Warning message
Global Const MB_ICONINFORMATION = 64   ' Information message

Global Const MB_APPLMODAL = 0          ' Application Modal Message Box
Global Const MB_DEFBUTTON1 = 0         ' First button is default
Global Const MB_DEFBUTTON2 = 256       ' Second button is default
Global Const MB_DEFBUTTON3 = 512       ' Third button is default
Global Const MB_SYSTEMMODAL = 4096      'System Modal
' Colors
Global Const BLACK = &H0&
Global Const Red = &HFF&
Global Const Green = &HFF00&
Global Const YELLOW = &HFFFF&
Global Const Blue = &HFF0000
Global Const MAGENTA = &HFF00FF
Global Const CYAN = &HFFFF00
Global Const WHITE = &HFFFFFF
' MsgBox return values
Global Const IDOK = 1                  ' OK button pressed
Global Const IDCANCEL = 2              ' Cancel button pressed
Global Const IDABORT = 3               ' Abort button pressed
Global Const IDRETRY = 4               ' Retry button pressed
Global Const IDIGNORE = 5              ' Ignore button pressed
Global Const IDYES = 6                 ' Yes button pressed
Global Const IDNO = 7                  ' No button pressed


'SendMessage constants.
Public Const WM_NCACTIVATE  As Long = &H86
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOW = 5
Public Function MakeNewRecordset(rs As Recordset, strOldLanguage As String, strNewLanguage As String)
Dim rsClone As Recordset, fld As Field, n As Integer
    On Error GoTo errNewRecordset
    Set rsClone = rs.Clone()
    With rsClone
        .MoveLast
        .MoveFirst
        For i = 0 To .RecordCount - 1
            If .Fields("Language") = strOldLanguage Then
                rs.AddNew
                rs.Fields("Language") = Trim(strNewLanguage)
                For n = 1 To rsClone.Fields.Count - 1
                    rs.Fields(n) = rsClone.Fields(n)
                Next
                rs.Update
                Exit Function
            End If
            .MoveNext
        Next
    End With
    Exit Function
    
errNewRecordset:
    Beep
    MsgBox Err.Description, vbCritical, "New Recordset"
    Err.Clear
End Function

Public Function MakeNewLanguage(strLanguage As String)
Dim rs As Recordset, iNo As Integer, boolNotFound As Boolean, iCount As Integer
Dim db As DAO.Database
Dim tbl As DAO.TableDef
    Set db = DBEngine.OpenDatabase(App.Path & "\Recipe.mdb")
    iCount = 0
    On Error GoTo errNewLanguage
    For Each tbl In db.TableDefs
        Select Case tbl.Name
        Case "MSysAccessObjects"
        Case "MSysObjects"
        Case "MSysQueries"
        Case "MSysRelationships"
        Case "MSysACEs"
        Case "SpellLanguage"
        Case "frmAbout", "frmCountry", "frmDimension", "frmHaveIngredients", "frmIngredients", "frmLinks", "frmMail", "frmRecipe", "frmScreenLanguage", "frmUser", "frmWriteToMe", "frmRecipeTypes"
            Set rs = db.OpenRecordset(tbl.Name)
            rs.MoveLast
            iNo = rs.RecordCount
            iCount = iCount + 1
            rs.MoveFirst
            boolNotFound = True
            For i = 0 To iNo - 1
                If Trim(rs.Fields("Language")) = Trim(strLanguage) Then
                    boolNotFound = False
                    Exit For
                End If
            rs.MoveNext
            Next
            If boolNotFound Then
                Call MakeNewRecordset(rs, "ENG", strLanguage)
            End If
        Case Else
        End Select
    Next
    Exit Function
    
errNewLanguage:
    Beep
    MsgBox Err.Description, vbInformation, "New Language"
    Resume Next
End Function

Public Function IsLanguagePresent(rsLanguage As Recordset, strLanguage As String) As Boolean
    IsLanguagePresent = False
    On Error GoTo errLangPres
    With rsLanguage
        .MoveLast
        .MoveFirst
        For i = 0 To .RecordCount - 1
            If .Fields("Language") = strLanguage Then
                IsLanguagePresent = True
                Exit For
            End If
        .MoveNext
        Next
    End With
    Exit Function
    
errLangPres:
    Beep
    MsgBox Err.Description, vbExclamation, "Is Language Present"
    Err.Clear
End Function

Public Function getstring(hKey As Long, strPath As String, strValue As String)
    'EXAMPLE:
    'text1.text = getstring(HKEY_CURRENT_USER, "Software\VBW\Registry", "String")
    Dim R
    Dim lValueType
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    
    R = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                getstring = Left$(strBuf, intZeroPos - 1)
            Else
                getstring = strBuf
            End If
        End If
    End If
End Function

Public Sub PrintLongText(sText As String, X As Integer)
'Public method to print text
    Dim i As Integer, J As Integer, sCurrWord As String
    
    m_LeftMargin = X
    Printer.ScaleMode = vbMillimeters
    
    'Print text, word-wrapping as we go
    i = 1
    Do Until i > Len(sText)
        'Get next word
        sCurrWord = ""
        Do Until i > Len(sText) Or Mid$(sText, i, 1) <= " "
            sCurrWord = sCurrWord & Mid$(sText, i, 1)
            i = i + 1
        Loop
        'Check if word will fit on this line
        If (Printer.CurrentX + Printer.TextWidth(sCurrWord)) > (Printer.ScaleWidth - m_RightMargin) Then
            'Send carriage-return line-feed to printer
            Printer.Print
            'Check if we need to start a new page
            If Printer.CurrentY > (Printer.ScaleHeight - m_BottomMargin) Then
                Printer.NewPage
                Printer.CurrentY = m_TopMargin
            Else
                Printer.CurrentX = m_LeftMargin
            End If
        End If
        
        'Print this word
        Printer.Print sCurrWord;
        'Process whitespace and any control characters
        Do Until i > Len(sText) Or Mid$(sText, i, 1) > " "
            Select Case Mid$(sText, i, 1)
                Case " "        'Space
                    Printer.Print " ";
                Case Chr$(10)   'Line-feed
                    'Send carriage-return line-feed to printer
                    Printer.Print
                    'Check if we need to start a new page
                    If Printer.CurrentY > (Printer.ScaleHeight - m_BottomMargin) Then
                        Printer.NewPage
                        Printer.CurrentY = m_TopMargin
                    Else
                        Printer.CurrentX = m_LeftMargin
                    End If
                Case Chr$(9)    'Tab
                    J = (Printer.CurrentX - m_LeftMargin) / Printer.TextWidth("0")
                    J = J + (10 - (J Mod 10))
                    Printer.CurrentX = m_LeftMargin + (J * Printer.TextWidth("0"))
                Case Else       'Ignore other characters
            End Select
            i = i + 1
        Loop
    Loop
    'Printer.EndDoc
End Sub

Public Function IsWebConnected(Optional ByRef ConnType As String) As Boolean
Dim dwFlags As Long
Dim WebTest As Boolean
    ConnType = ""
    WebTest = InternetGetConnectedState(dwFlags, 0&)
    Select Case WebTest
        Case dwFlags And CONNECT_LAN: ConnType = "LAN"
        Case dwFlags And CONNECT_MODEM: ConnType = "Modem"
        Case dwFlags And CONNECT_PROXY: ConnType = "Proxy"
        Case dwFlags And CONNECT_OFFLINE: ConnType = "Offline"
        Case dwFlags And CONNECT_CONFIGURED: ConnType = "Configured"
        'Case dwflags And CONNECT_RAS:
        'ConnType = "Remote"
    End Select
    IsWebConnected = WebTest
End Function

Public Function PrintPictureToFitPage(Prn As Printer, pic As Picture) As Boolean
Const vbHiMetric As Integer = 8
Dim PicRatio      As Double
Dim PrnWidth      As Double
Dim PrnHeight     As Double
Dim PrnRatio      As Double
Dim PrnPicWidth   As Double
Dim PrnPicHeight  As Double

    On Error GoTo ErrorHandler
    
    ' *** Determine if picture should be printed in
    'landscape or portrait and set the orientation
    If pic.Height >= pic.Width Then
       Prn.Orientation = vbPRORPortrait ' Taller than wide
    Else
       Prn.Orientation = vbPRORLandscape ' Wider than tall
    End If
    
    ' *** Calculate device independent Width to Height ratio for picture
    PicRatio = pic.Width / pic.Height
    
    ' *** Calculate the dimentions of the printable area in HiMetric
    PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
    PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
    
    ' *** Calculate device independent Width to Height ratio for printer
    PrnRatio = PrnWidth / PrnHeight
    
    ' *** Scale the output to the printable area
    If PicRatio >= PrnRatio Then
       ' *** Scale picture to fit full width of printable area
       PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, _
           Prn.ScaleMode)
       PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, _
           vbHiMetric, Prn.ScaleMode)
    Else
       ' *** Scale picture to fit full height of printable area
       PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, _
           Prn.ScaleMode)
       PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, _
           vbHiMetric, Prn.ScaleMode)
    End If
    
    ' *** Print the picture using the PaintPicture method
    Prn.PaintPicture pic, 0, 0, PrnPicWidth, PrnPicHeight
    Prn.EndDoc
    PrintPictureToFitPage = True
    Exit Function

ErrorHandler:
    PrintPictureToFitPage = False
End Function



Sub Dither(vForm As Form)
Dim intLoop As Integer
    vForm.AutoRedraw = True
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub

Public Sub SendOutlookMail(Subject As String, Recipient As String, Message As String)
    On Error GoTo ErrorHandler
    Dim oLapp As Object
    Dim oItem As Object
    
    Set oLapp = CreateObject("Outlook.application")
    Set oItem = oLapp.CreateItem(0)
   
    With oItem
       .Subject = Subject
       .To = Recipient
       .Body = Message
       .Send
    End With
    
ErrorHandler:
    Set oLapp = Nothing
    Set oItem = Nothing
    Exit Sub
End Sub

Public Sub onGotFocus()
    If TypeOf Screen.ActiveControl Is TextBox Then
        With Screen.ActiveControl
            .SelStart = 0
            .SelLength = Len(.Text)
            .ForeColor = vbWindowText
        End With
    End If
End Sub


'this sub is used for writing all errors to a common log file
Public Sub WriteLogFile(Frm As String, ssString As String)
    Open "MasterHom.log" For Append As #1
        Write #1, Now, "IPB System", Frm, ssString
    Close #1
End Sub


