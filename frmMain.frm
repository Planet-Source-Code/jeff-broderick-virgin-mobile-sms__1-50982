VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "VMobile - [QuickMessenger]"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameSent 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   0
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Image cmdExit 
         Height          =   375
         Left            =   2880
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Image cmdSendAnother 
         Height          =   375
         Left            =   840
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   840
         Picture         =   "frmMain.frx":3072
         Top             =   3240
         Width           =   3180
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Your message is on its way!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   1440
         Width           =   3375
      End
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      Height          =   2055
      Left            =   120
      MaxLength       =   125
      TabIndex        =   6
      Top             =   3000
      Width           =   4695
   End
   Begin VB.TextBox txtFrom 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   4695
   End
   Begin VB.TextBox txtTo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4695
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   5535
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   5055
      ExtentX         =   8916
      ExtentY         =   9763
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label lblMoreInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "More Information"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image cmdSend 
      Height          =   375
      Left            =   3360
      Picture         =   "frmMain.frx":6C54
      Top             =   5160
      Width           =   1500
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "© Virgin Mobile USA, LLC 2002-2003. All Rights Reserved."
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   5640
      Width           =   4935
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "QuickMessenger"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1020
      Left            =   240
      Picture         =   "frmMain.frx":89E2
      Top             =   0
      Width           =   4320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Message: (max characters: 125) "
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label lblFrom 
      BackStyle       =   0  'Transparent
      Caption         =   "From: (you) "
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label lblTo 
      BackStyle       =   0  'Transparent
      Caption         =   "To: (mobile phone number) "
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000100CC&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1035
      Left            =   0
      Top             =   0
      Width           =   5055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000A9&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   0
      Top             =   840
      Width           =   5055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdSend_Click()
    Dim oldstring As String, newletter As String, oldletter As String, newstring As String
    If Me.txtTo.Text = "" Then MsgBox "Missing Information: To:"
    If Me.txtFrom.Text = "" Then MsgBox "Missing Information: From:"
    If Me.txtMessage.Text = "" Then MsgBox "Missing Information: Message:"
    oldstring = Me.txtMessage.Text
    newletter = "+"
    oldletter = " "
    newstring = Replace(oldstring, newletter, oldletter)
    Me.wb.Navigate "http://www.virginmobileusa.com/xtras/messaging/processSMSMessage.do?to=" & Me.txtTo.Text & "&from=" & Me.txtFrom & "&message=" & newstring
    Me.frameSent.Visible = True
End Sub

Public Function Replace(oldstring, newletter, oldletter) As String
    Dim i As Integer
    i = 1


    Do While InStr(i, oldstring, oldletter, vbTextCompare) <> 0
        Replace = Replace & Mid(oldstring, i, InStr(i, oldstring, oldletter, vbTextCompare) - i) & newletter
        i = InStr(i, oldstring, oldletter, vbTextCompare) + Len(oldletter)
    Loop
    Replace = Replace & Right(oldstring, Len(oldstring) - i + 1)
End Function

Private Sub cmdSendAnother_Click()
Me.txtTo.Text = ""
Me.txtFrom.Text = ""
Me.txtMessage.Text = ""
Me.frameSent.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lblMoreInfo_Click()
frmInfo.Show 1
End Sub
