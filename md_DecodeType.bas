Attribute VB_Name = "md_DecodeType"
Option Explicit

Type pxVL
    Header As Integer
    Length As Integer
End Type
'0069 : account information & charserver list
Type px0069
    Header As Integer
    Length As Integer
    SessionID As Long
    AccountID As Long
    Token As Long
    ServerInfo As String * 30
    Sex As Byte
End Type
' server_info
'   00 - normal
'   01 - under maintenance
'   02 - for 18+ only
'   04 - free
'   03-ff - reserved
Type px0069ex
    IP As String * 4
    Port As Integer
    Name As String * 20
    Players As Integer
    ServInfo As Integer
    Newx As Integer
End Type

Type px006Bex
    CharID As Long
    expBASE As Long
    Zeny As Long
    expJOB As Long
    levelJOB As Long
    CharState As Long
    Aliment As Long
    Options As Long
    Karma As Long
    Manner As Long
    
    StatusPoint As Integer
    
    HP As Integer
    HPmax As Integer
    SP As Integer
    SPmax As Integer
    walkSpeed As Integer
    jobID As Integer
    HairStyle As Integer
    Weapon As Integer
    
    levelBase As Integer
    SkillPoint As Integer
    
    UnProcess As String * 12
    Name As String * 24
    STR As Byte
    AGI As Byte
    VIT As Byte
    INT As Byte
    DEX As Byte
    LUK As Byte
    Index As Byte
    Reserved As Byte
End Type
'00A3: list of consumptive item and collecter item that you have
'{<index>.w <item ID>.w <type>.B <identify flag>.B <amount>.w ?.2B}.10B*
Type px00A3ex
    Index As Integer
    ItemID As Integer
    ItemType As Byte
    Identified As Byte
    Amount As Integer
    Unknown As Integer
End Type
'00A4: list of equipments that you have
Type px00A4ex
    Index As Integer
    ItemID As Integer
    ItemType As Byte
    Identified As Byte
    EqType As Integer
    EqPlave As Integer
    Attribute As Byte
    Refine As Byte
    Note As String * 8
End Type
'010F: list skills
'{<skill ID>.w <target type>.w ?.w <lv>.w <sp>.w <range>.w <skill name>.24B <up>.B}.37B*
Type px010Fex
    SkillID As Integer
    Target As Integer
    Unknown As Integer
    Level As Integer
    SP As Integer
    Range As Integer
    SkillName As String * 24
    CanUP As Byte
End Type
'0122: List Equipments in cart
Type px0122ex
    Index As Integer '2
    ItemID As Integer '2
    ItemType As Byte '1
    Identified As Byte '1
    EqType As Integer '2
    EqPlace As Integer '2
    Attribute As Byte '1
    Refine As Byte '1
    Note As String * 8 '8
End Type
Type px0122exx
    Flag As Integer
    Element As Byte
    Strength As Byte
    BSID As Long
End Type
'0123: List items in cart
'{<index>.w <item ID>.w <type>.B <identify flag>.B <amount>.w ?.2B}.10B*
Type px0123ex
    Index As Integer
    ItemID As Integer
    ItemType As Byte
    Identified As Byte
    Amount As Integer
    Unknown As Integer
End Type
