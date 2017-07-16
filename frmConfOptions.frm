VERSION 5.00
Begin VB.Form frmConfOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Config options.ini"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   5775
   Begin VB.Frame frmMap 
      Caption         =   "Map Control Options"
      Height          =   3255
      Left            =   1800
      TabIndex        =   2
      Top             =   0
      Width           =   3975
      Begin VB.CheckBox chkMap 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Section"
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.ListBox lstSection 
         Appearance      =   0  'Flat
         Height          =   2760
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmConfOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    PrepareSection
End Sub

Private Sub PrepareSection()
    With lstSection
        .Clear
        .AddItem "Map Control"
        .AddItem "Startup Mode"
        .AddItem "AI Control"
        .AddItem "Skill Use"
        .AddItem "Monk"
        .AddItem "HP/SP"
        .AddItem "Item Using"
        .AddItem "Tele/DC"
        .AddItem "Pet Control"
        .AddItem "Avoid Control"
        .AddItem "Attack Control"
        .AddItem "Range Control"
        .AddItem "Weight Control"
        .AddItem "Timing Control"
    End With
End Sub
