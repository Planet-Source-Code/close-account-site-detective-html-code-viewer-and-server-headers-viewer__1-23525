VERSION 5.00
Begin VB.Form frmProxy 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proxy Configuration"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   1950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtPort 
      BackColor       =   &H00C0C000&
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "8080"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtProxy 
      BackColor       =   &H00C0C000&
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblPort 
      BackColor       =   &H000080FF&
      Caption         =   "Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblProxy 
      BackColor       =   &H000080FF&
      Caption         =   "Proxy:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Another one of my old codes - Sent to planetsourcecode on May 29 2001
'by David Fial - djf1010@aol.com

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim boolNumeric As Boolean
    boolNumeric = IsNumeric(txtPort.Text)
    If boolNumeric = False Then
        MsgBox "The port you specified is NOT an integer. The default port for a proxy is 8080", , "Invalid port."
        Exit Sub
    End If
    strProxy = txtProxy.Text
    intPort = txtPort.Text
    intPortProx = txtPort.Text
    Unload Me
End Sub

Private Sub Form_Load()
    If strProxy <> "" Then
        txtProxy.Text = strProxy
        txtPort.Text = intPortProx
    End If
End Sub
