VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Site Detective by David Fiala"
   ClientHeight    =   7965
   ClientLeft      =   870
   ClientTop       =   330
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   3120
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "sd"
      DialogTitle     =   "Site Detective"
      Filter          =   "*.sd"
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Settings"
      Height          =   255
      Left            =   4320
      TabIndex        =   25
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Settings"
      Height          =   255
      Left            =   4320
      TabIndex        =   24
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdProxy 
      Caption         =   "Config Proxy"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CheckBox chkProxy 
      BackColor       =   &H00FF8080&
      Caption         =   "Use Proxy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1560
      TabIndex        =   10
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdDoIt 
      Caption         =   "Execute"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Frame fraForm 
      BackColor       =   &H00FF8080&
      Caption         =   "Form Data"
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   120
      TabIndex        =   21
      Top             =   4560
      Width           =   6135
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove Selected"
         Height          =   255
         Left            =   4440
         TabIndex        =   26
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox cmbMethod 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear All"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox cmbFormData 
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   1680
         Left            =   1680
         Style           =   1  'Simple Combo
         TabIndex        =   7
         Top             =   600
         Width           =   4215
      End
      Begin VB.CheckBox chkForm 
         BackColor       =   &H000000FF&
         Caption         =   "*IMPORTANT* CHECK THIS BOX IF YOU WANT TO USE FORM DATA."
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label lblFormExample 
         BackColor       =   &H00FF8080&
         Caption         =   "EG: sex=male"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblFormInfo 
         BackColor       =   &H00FF8080&
         Caption         =   "To add fields to the form put them in the combo box in the format of: NAME=VALUE"
         Height          =   975
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame fraUserAgent 
      BackColor       =   &H00FF8080&
      Caption         =   "User Agent?"
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   6135
      Begin VB.ListBox lstUserAgent 
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   660
         ItemData        =   "frmMain.frx":0019
         Left            =   3240
         List            =   "frmMain.frx":0023
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblUserAgentInfo 
         BackColor       =   &H00FF8080&
         Caption         =   "Select an OS you want to become for the User Agent:"
         Height          =   435
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   2130
      End
   End
   Begin VB.Frame fraCookies 
      BackColor       =   &H00FF8080&
      Caption         =   "Cookies?"
      ForeColor       =   &H00020002&
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   6135
      Begin VB.TextBox txtCookies 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label lblCookies 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Cookie(s):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblCookieInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "If you don't want to send the server cookies, then leave blank below empty."
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame fraSite 
      BackColor       =   &H00FF8080&
      Caption         =   "What site?"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   6135
      Begin VB.TextBox txtFile 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Text            =   "index.html"
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   360
         Left            =   840
         TabIndex        =   0
         Text            =   "www.msn.com"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblSlash 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2520
         TabIndex        =   15
         Top             =   240
         Width           =   75
      End
      Begin VB.Label lblHttp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "http://"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "djf1010@aol.com"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4800
      TabIndex        =   27
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Line linDivider 
      BorderColor     =   &H0000FF00&
      X1              =   120
      X2              =   6240
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblAppName 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Site Detective"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Another one of my old codes - Sent to planetsourcecode on May 29 2001
'by David Fial - djf1010@aol.com

Option Explicit

Public Sub LoadFormData(strData As String)
    cmbFormData.Clear
    Do Until Len(strData) <= 1
        cmbFormData.AddItem (Mid(strData, 1, InStr(1, strData, "&") - 1))
        strData = Mid(strData, InStr(1, strData, "&") + 1)
    Loop
End Sub

Private Sub chkProxy_Click()
    If chkProxy.Value = 1 Then
        cmdProxy.Enabled = True
    Else
        cmdProxy.Enabled = False
    End If
End Sub

Private Sub cmbFormData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Len(cmbFormData.Text) < 3 Then Exit Sub
        If InStr(1, cmbFormData.Text, "=", vbTextCompare) = 0 Then
            MsgBox "You didnt put a * character in the box." & vbCrLf & "Valid format is: NAME=VALUE" & vbCrLf & vbCrLf & "EG: sex=male" & vbCrLf & "EG: sex=female"
            Exit Sub
        End If
        If cmbFormData.Text = "" Then Exit Sub
        cmbFormData.Text = Replace(cmbFormData.Text, Chr(32), "%20")
        cmbFormData.AddItem cmbFormData.Text
        cmbFormData.Text = ""
    End If
End Sub

Private Sub cmdClear_Click()
    Dim intConfirmClear As Integer
    intConfirmClear = MsgBox("Are you sure you want to clear the form data?", vbYesNo + vbExclamation + vbDefaultButton2, "Really clear?")
    If intConfirmClear = vbYes Then cmbFormData.Clear
End Sub

Private Sub cmdDoIt_Click()
    Dim strFormData As String
    strHeaders = ""
    If chkProxy.Value = 1 Then
        strServer = strProxy
        intPort = intPortProx
    Else
        strServer = txtServer.Text
        intPort = 80
    End If
    If chkForm.Value = 1 Then
        Dim intTemp As Integer
        For intTemp = 0 To cmbFormData.ListCount Step 1
            strFormData = strFormData & cmbFormData.List(intTemp) & "&"
    
        Next
        strFormData = Mid(strFormData, 1, Len(strFormData) - 2)
    End If
    If chkForm.Value = 1 And cmbMethod.Text = "POST" Then
        'POST
        If chkProxy.Value = 1 Then
            strHeaders = "POST http://" & txtServer.Text & "/" & txtFile.Text & " HTTP/1.0" & vbCrLf
        Else
            strHeaders = "POST /" & txtFile.Text & " HTTP/1.0" & vbCrLf
        End If
        strHeaders = strHeaders & "Accept: */*" & vbCrLf
        strHeaders = strHeaders & "Accept-Language: en-us" & vbCrLf
        strHeaders = strHeaders & "Content-Encoding: gzip, deflate" & vbCrLf
        strHeaders = strHeaders & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
        If txtCookies.Text <> "" Then strHeaders = strHeaders & "Cookie: " & txtCookies.Text & vbCrLf
        strHeaders = strHeaders & "Host: " & txtServer.Text & vbCrLf
        strHeaders = strHeaders & "User-Agent: Mozilla 4.0 (" & lstUserAgent.Text & ")" & vbCrLf
        strHeaders = strHeaders & "Content-Length:" & Len(strFormData) & vbCrLf
        If chkProxy.Value = 1 Then
            strHeaders = strHeaders & "Proxy-Connection: Close" & vbCrLf & vbCrLf
        Else
            strHeaders = strHeaders & "Connection: Close" & vbCrLf & vbCrLf
        End If
        strHeaders = strHeaders & strFormData
    Else
        'GET
        If chkProxy.Value = 1 Then
            If chkForm.Value = 1 Then strHeaders = "GET http://" & txtServer.Text & "/" & txtFile.Text & "?" & strFormData & " HTTP/1.0" & vbCrLf
            If chkForm.Value = 0 Then strHeaders = "GET http://" & txtServer.Text & "/" & txtFile.Text & " HTTP/1.0" & vbCrLf
        Else
            If chkForm.Value = 1 Then strHeaders = "GET /" & txtFile.Text & "?" & strFormData & " HTTP/1.0" & vbCrLf
            If chkForm.Value = 0 Then strHeaders = "GET /" & txtFile.Text & " HTTP/1.0" & vbCrLf
        End If
        strHeaders = strHeaders & "Accept: */*" & vbCrLf
        strHeaders = strHeaders & "Accept-Language: en-us" & vbCrLf
        strHeaders = strHeaders & "Content-Encoding: gzip, deflate" & vbCrLf
        If txtCookies.Text <> "" Then strHeaders = strHeaders & "Cookie: " & txtCookies.Text & vbCrLf
        strHeaders = strHeaders & "Host: " & txtServer.Text & vbCrLf
        strHeaders = strHeaders & "User-Agent: Mozilla 4.0 (" & lstUserAgent.Text & ")" & vbCrLf
        If chkProxy.Value = 1 Then
            strHeaders = strHeaders & "Proxy-Connection: Close" & vbCrLf & vbCrLf
        Else
            strHeaders = strHeaders & "Connection: Close" & vbCrLf & vbCrLf
        End If
    End If
    'MsgBox strHeaders
    'Exit Sub
    frmRaw.Show vbModal
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLoad_Click()
    Dim intConfirm As Integer
    intConfirm = MsgBox("WARNING: If you choose to continue all current settings will be lost unless they are saved!" & vbCrLf & vbCrLf & "Continue?", vbCritical + vbApplicationModal + vbYesNo, "***WARNING***")
    If intConfirm = vbNo Then Exit Sub
    strSettingFile = ""
    cdMain.filename = ""
    cdMain.ShowOpen
    strSettingFile = cdMain.filename
    If strSettingFile = "" Then
        MsgBox "ERROR: No file was selected,invalid file, or you pressed cancel. Load aborted."
        Exit Sub
    End If
    txtServer.Text = ReadSet("txtServer")
    txtFile.Text = ReadSet("txtFile")
    txtCookies.Text = ReadSet("txtCookies")
    lstUserAgent.Text = ReadSet("lstUserAgent")
    chkForm.Value = ReadSet("chkForm")
    cmbMethod.Text = ReadSet("cmbMethod")
    LoadFormData (ReadSet("cmbFormData"))
    chkProxy.Value = ReadSet("chkProxy")
    strProxy = ReadSet("strProxy")
    intPortProx = CInt(ReadSet("intPortProx"))
    intPort = CInt(ReadSet("intPort"))
    MsgBox "Settings loaded from: " & strSettingFile
End Sub

Private Sub cmdProxy_Click()
    frmProxy.Show vbModal
End Sub

Private Sub cmdRemove_Click()
    Dim intConfirm As Integer
    intConfirm = MsgBox("Do you really want to remove: " & cmbFormData.Text & " from the form?", vbYesNo + vbExclamation + vbDefaultButton2, "About to remove...?")
    If intConfirm = vbNo Then Exit Sub
    Dim intTemp As Integer
    For intTemp = 0 To cmbFormData.ListCount Step 1
        If cmbFormData.List(intTemp) = cmbFormData.Text Then
            cmbFormData.RemoveItem (intTemp)
            Exit For
        End If
    Next
End Sub

Private Sub cmdSave_Click()
    Dim intTemp As Integer
    Dim strFormData As String
    strSettingFile = ""
    cdMain.filename = ""
    cdMain.ShowSave
    strSettingFile = cdMain.filename
    If strSettingFile = "" Then
        MsgBox "ERROR: No file was selected,invalid file, or you pressed cancel. Save aborted."
        Exit Sub
    End If
    
    For intTemp = 0 To cmbFormData.ListCount Step 1
        strFormData = strFormData & cmbFormData.List(intTemp) & "&"
    Next
    If strFormData <> "" Then
        strFormData = Mid(strFormData, 1, Len(strFormData) - 2)
    End If
    WriteSet "txtServer", txtServer.Text
    WriteSet "txtFile", txtFile.Text
    WriteSet "txtCookies", txtCookies.Text
    WriteSet "lstUserAgent", lstUserAgent.Text
    WriteSet "chkForm", chkForm.Value
    WriteSet "cmbMethod", cmbMethod.Text
    WriteSet "cmbFormData", strFormData & "&"
    WriteSet "chkProxy", chkProxy.Value
    WriteSet "strProxy", strProxy
    WriteSet "intPortProx", intPortProx
    WriteSet "intPort", intPort
    
    MsgBox "Complete. Saved to: " & strSettingFile
End Sub

Private Sub Form_Load()
    cdMain.InitDir = App.Path
    cmbMethod.Text = "GET"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim intConfirmExit As Integer
    intConfirmExit = MsgBox("Do you really want to exit Site Detective?", vbSystemModal + vbYesNo + vbDefaultButton2, "Really exit?")
    If intConfirmExit = vbNo Then Cancel = True
End Sub
