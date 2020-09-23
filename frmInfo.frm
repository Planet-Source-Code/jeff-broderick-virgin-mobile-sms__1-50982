VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "More Information"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   3615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmInfo.frx":3072
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "© Virgin Mobile USA, LLC 2002-2003. All Rights Reserved."
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5520
      Width           =   4935
   End
   Begin VB.Label lblSite 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.virginmobile.com/"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Can't talk? Then txt instead. "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1020
      Left            =   240
      Picture         =   "frmInfo.frx":3342
      Top             =   0
      Width           =   4320
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
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblSite_Click()
Call OpenBrowser("http://www.virginmobile.com/", frmInfo.hwnd)
End Sub

Public Function OpenBrowser(strURL As String, lngHwnd As Long)
    OpenBrowser = ShellExecute(lngHwnd, vbNullString, strURL, vbNullString, "c:\", SW_SHOWDEFAULT)
End Function
