VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRaw 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Site Detective"
   ClientHeight    =   10275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   14550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find It(Down)"
      Height          =   375
      Left            =   12960
      TabIndex        =   4
      Top             =   9480
      Width           =   1455
   End
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   3
      Top             =   9480
      Width           =   3375
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   10020
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Hold on..."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckRaw 
      Left            =   1800
      Top             =   9480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.CommandButton cmdLeave 
      Caption         =   "Done(Leave)"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   9480
      Width           =   1455
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Height          =   9375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   14535
   End
End
Attribute VB_Name = "frmRaw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Another one of my old codes - Sent to planetsourcecode on May 29 2001
'by David Fial - djf1010@aol.com

Option Explicit

Private Sub cmdFind_Click()
    On Error GoTo errorh
    txtData.SelStart = InStr(txtData.SelStart, txtData.Text, txtFind.Text, vbTextCompare)
    txtData.SelLength = Len(txtFind.Text)
    Exit Sub
errorh:
    MsgBox "An error has occured: " & Err.Description
    Resume Next
End Sub

Private Sub cmdLeave_Click()
Unload Me
End Sub

Private Sub Form_Load()
    sckRaw.RemoteHost = strServer
    sckRaw.RemotePort = intPort
    sckRaw.Connect
    StatusBar.SimpleText = "Connecting to: " & sckRaw.RemoteHost
End Sub

Private Sub sckRaw_Close()
    StatusBar.SimpleText = "Disconnected from: " & sckRaw.RemoteHost
End Sub

Private Sub sckRaw_Connect()
    sckRaw.SendData strHeaders
    StatusBar.SimpleText = "Connected to: " & sckRaw.RemoteHost
End Sub

Private Sub sckRaw_DataArrival(ByVal bytesTotal As Long)
    Dim strTempData As String
    sckRaw.GetData strTempData
    txtData.Text = txtData.Text & strTempData
    StatusBar.SimpleText = "Getting data from: " & sckRaw.RemoteHost
End Sub

Private Sub sckRaw_SendComplete()
    StatusBar.SimpleText = "Sending request to: " & sckRaw.RemoteHost
End Sub
