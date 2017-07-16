VERSION 5.00
Begin VB.Form frmPopupChat 
   Caption         =   "Popup Menu"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuPet 
      Caption         =   "Pet"
      Begin VB.Menu mnuFeed 
         Caption         =   "Feeds"
      End
      Begin VB.Menu mnuPerformance 
         Caption         =   "Performance"
      End
      Begin VB.Menu mnuEgg 
         Caption         =   "Back to Egg!"
      End
   End
   Begin VB.Menu mnuItem 
      Caption         =   "Use Item"
      Begin VB.Menu mnuUEquip 
         Caption         =   "Un-Equip"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEquip 
         Caption         =   "Equip"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUse 
         Caption         =   "Use"
      End
      Begin VB.Menu mnuDrop 
         Caption         =   "Drop"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCart 
         Caption         =   "Move to cart"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuChat 
      Caption         =   "Chat Select"
      Begin VB.Menu mnuPublic 
         Caption         =   "Public"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuWhisper 
         Caption         =   "Whisper"
      End
      Begin VB.Menu mnuParty 
         Caption         =   "Party"
      End
      Begin VB.Menu mnuGuild 
         Caption         =   "Guild"
      End
   End
   Begin VB.Menu mnuBuy 
      Caption         =   "Buying"
      Begin VB.Menu mnuBuyAdd 
         Caption         =   "Add new"
      End
      Begin VB.Menu mnuBuyDelete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuPartylist 
      Caption         =   "PartyList"
      Visible         =   0   'False
      Begin VB.Menu mnuPartys 
         Caption         =   "-Party-"
      End
      Begin VB.Menu mnuAddParty 
         Caption         =   "Add"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuKickParty 
         Caption         =   "Kick"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLeaveParty 
         Caption         =   "Leave"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGuilds 
         Caption         =   "-Guild-"
      End
      Begin VB.Menu mnuAddGuild 
         Caption         =   "Add"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuKickGuild 
         Caption         =   "Kick"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLeaveGuild 
         Caption         =   "Leave"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmPopupChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuBuyAdd_Click()
    'ReDim Preserve ItemCtrl(uboud(ItemCtrl) + 1)
    'update_buyoption
End Sub

Private Sub mnuBuyDelete_Click()
    'tmp
End Sub

Private Sub mnuCart_Click()
    'frmmain.pkt_CartGet
    Dim Index  As Long, Amount As Long
    Index = Val(frmItem.lstInvent.List(frmItem.lstInvent.ListIndex))
    If AllInv(Index).Amount > 1 Then
        Amount = Val(InputBox("Enter number to add", "Get from inventory", CStr(AllInv(Index).Amount)))
        If Amount > AllInv(Index).Amount Then Amount = AllInv(Index).Amount
    Else
        Amount = 1
    End If
    If Amount > 0 Then
        pkt_CartGet Index, Amount
        Stat "Cart item from inventory : "
        Stat "[" & AllInv(Index).Name & "]", vbBlue
        Stat " " & Amount & "EA" + vbCrLf
    End If
End Sub

Private Sub mnuDrop_Click()
    frmMain.Drop_Item
End Sub

Private Sub mnuEgg_Click()
frmMain.BackEgg
End Sub

Private Sub mnuEquip_Click()
    frmMain.Equip_Item
End Sub

Private Sub mnuFeed_Click()
frmMain.PetFeed
End Sub

Public Sub mnuGuild_Click()
    mnuPublic.CheckED = False
    mnuWhisper.CheckED = False
    mnuGuild.CheckED = True
    mnuParty.CheckED = False
End Sub

'Private Sub mnuInfo_Click()
'With frmItem.lstInvent
'    If .List(.ListIndex) <> "" Then
'        frmDescription.LabName.Caption = Return_ItemName(AllInv(Val(.List(.ListIndex))).Name)
'        frmDescription.Text1.text = Return_ItemD(AllInv(Val(.List(.ListIndex))).NameID)
'        frmDescription.Visible = True
'    End If
'End With
'End Sub

Public Sub mnuParty_Click()
    mnuPublic.CheckED = False
    mnuWhisper.CheckED = False
    mnuGuild.CheckED = False
    mnuParty.CheckED = True
End Sub

Private Sub mnuPerformance_Click()
frmMain.PetPerformance
End Sub

Public Sub mnuPublic_Click()
    mnuPublic.CheckED = True
    mnuWhisper.CheckED = False
     mnuGuild.CheckED = False
    mnuParty.CheckED = False
End Sub

Private Sub mnuUEquip_Click()
    frmMain.unEquip_Item
End Sub

Private Sub mnuUse_Click()
    frmMain.Use_Item
End Sub

Public Sub mnuWhisper_Click()
    mnuPublic.CheckED = False
    mnuWhisper.CheckED = True
     mnuGuild.CheckED = False
    mnuParty.CheckED = False
End Sub
Private Sub mnuAddParty_Click()
    frmMain.Send_AddParty
End Sub

Private Sub mnuKickParty_Click()
    frmMain.Send_KickParty
End Sub

Private Sub mnuLeaveParty_Click()
    frmMain.Send_LeaveParty
End Sub

Private Sub mnuAddGuild_Click()
    frmMain.Send_AddGuild
End Sub

'Private Sub mnuKickGuild_Click()
'    frmMain.Send_KickGuild
'End Sub

'Private Sub mnuLeaveGuild_Click()
'    frmMain.Send_LeaveGuild
'End Sub

