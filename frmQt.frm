VERSION 5.00
Begin VB.Form frmQt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quantity ?"
   ClientHeight    =   660
   ClientLeft      =   8910
   ClientTop       =   9105
   ClientWidth     =   1935
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQt.frx":0000
   LinkTopic       =   "Quantity"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtQt 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAF9FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   1080
      Picture         =   "frmQt.frx":0E42
      Top             =   315
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   240
      Picture         =   "frmQt.frx":1380
      Top             =   315
      Width           =   630
   End
   Begin VB.Image Image3 
      Height          =   1545
      Left            =   -1200
      Picture         =   "frmQt.frx":1884
      Top             =   -880
      Width           =   4200
   End
End
Attribute VB_Name = "frmQt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
frmMain.Sell_Item
frmQt.Visible = False
End Sub

Private Sub Image2_Click()
frmQt.Visible = False
End Sub
