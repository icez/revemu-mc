// packet.h - RO EP 6.0 packet
// Collected by JunkBot
// last updated: 14 May 2004
//----------------------------------------------------------------------------

#ifndef _PACKET_H
#define _PACKET_H

#include <winsock.h>

#ifdef __cplusplus
extern "C" {
#endif

#pragma pack(1) // solve the alignment problem

//--------------------------------------------------------------------
// header for variable length message
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short msglen; // 02-03
} msghdr;

//--------------------------------------------------------------------
// 0x0000: Flushing ?? - we've got this message when connected to server
typedef struct
{
  unsigned short cmd;        // 00-01
  unsigned char  unknown[2]; // 02-03
} msg0x0000;

//--------------------------------------------------------------------
// 0x0064: Master Login
typedef struct
{
  unsigned short cmd;          // 00-01
  unsigned long  patch;        // 02-05
  unsigned char  username[24]; // 06-29
  unsigned char  passwd[24];   // 30-53
  unsigned char  version;      // 54
}  msg0x0064;

//--------------------------------------------------------------------
// 0x0065: Game Login
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned long  accountID; // 02-05
  unsigned long  sessionID; // 06-09
  unsigned long  token;     // 10-13
  unsigned short unknown;   // 14-15
  unsigned char  sex;       // 16
} msg0x0065;

//--------------------------------------------------------------------
// 0x0066: Character Login
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned char  charID; // 02
} msg0x0066;

//--------------------------------------------------------------------
// 0x0069: Char Server Information
typedef struct
{
  unsigned short cmd;            // 00-01
  unsigned short msglen;         // 02-03
  unsigned long  sessionID;      // 04-07
  unsigned long  accountID;      // 08-11
  unsigned long  token;          // 12-15
  unsigned char  serverinfo[30]; // 16-45
  unsigned char  sexID;        // 46
} msg0x0069;

// server_info
//   00 - normal
//   01 - under maintenance
//   02 - for 18+ only
//   03-ff - reserved

typedef struct
{
  struct in_addr ip;             // 00-03
  unsigned short port;           // 04-05
  unsigned char  servername[20]; // 06-25
  unsigned short players;        // 26-27
  unsigned short server_info;    // 28-29 - see above
  unsigned short new_server;     // 30-31 - if 1 - display "new" before server name
} msg0x0069ex;

//--------------------------------------------------------------------
// 0x006A: Login Error

// result:
//   00 - id unregistered
//   01 - incorrect password
//   02 - id expired
//   03 - unused
//   04 - blocked by GM
//   05 - client version mismatched
//   06 - blocked until %s (infor in errormsg)
//   07 - too many players are online
//   08 - ???
//   09-62 - reserved
//   63 - id has been deleted
//   64-ff - reserved

typedef struct
{
  unsigned short cmd;          // 00-01
  unsigned char  result;       // 02 - see above
  unsigned char  errormsg[20]; // 03-22
} msg0x006A;

//--------------------------------------------------------------------
// 0x006B: Character Information for Selection

// state:
//   00 - normal
//   01 - stoned (paralysed)
//   02 - frozen
//   03 - stun
//   04 - sleep
//   05 - reserved
//   06 - darkness

// ailments
//   0000 0000 0000 0001 - Poison
//   0000 0000 0000 0010 - Curse
//   0000 0000 0000 0100 - Silence
//   0000 0000 0000 1000 - Confusion
//   0000 0000 0001 0000 - blind
//   0000 0000 0010 0000 - something about pet?

// option
//   0000 0000 0000 0001 - Sight / Ruwach
//   0000 0000 0000 0010 - Hide
//   0000 0000 0000 0100 - Cloak
//   0000 0000 0000 1000 - Cart (Lv 1-40)
//   0000 0000 0001 0000 - Falcon
//   0000 0000 0010 0000 - Pecopeco
//   0000 0000 0100 0000 - Disappear
//   0000 0000 1000 0000 - Cart (Lv 41-65)
//   0000 0001 0000 0000 - Cart (Lv 66-80)
//   0000 0010 0000 0000 - Cart (Lv 81-90)
//   0000 0100 0000 0000 - Cart (Lv 91-99)
//   0000 1000 0000 0000 - Hideous Mark

// weapon
//   00 - Fist
//   01 - Dagger
//   02 - Sword
//   03 - 2-handed sword
//   04 - Spear
//   05 - 2-handed Spear
//   06 - Axe
//   07 - 2-handed Axe
//   08 - Mace
//   09 - Special Mace
//   10 - Wand
//   11 - Bow
//   12-15 reserved
//   16 - Katar
//   17 - Katar2
//   18 - Sword 2
//   19 - Axe 2
//   20 - Katar + Sword
//   21 - Katar + Axe
//   22 - Axe + Sword
//   23 - error

typedef struct
{
  unsigned long  charID;       // 000-003
  unsigned long  baseExp;      // 004-007
  unsigned long  zeny;         // 008-011
  unsigned long  jobExp;       // 012-015
  unsigned long  jobLv;        // 016-019

  unsigned long  state;        // 020-023 - see above
  unsigned long  ailments;     // 024-027 - see above
  unsigned long  options;      // 028-031 - see above
  unsigned long  karma;        // 032-035
  unsigned long  manner;       // 036-039

  unsigned short status_pt;    // 040-041 - unallocated status point

  unsigned short HP;           // 042-043
  unsigned short maxHP;        // 044-045
  unsigned short SP;           // 046-047
  unsigned short maxSP;        // 048-049
  unsigned short move_delay;   // 050-051
  unsigned short jobID;        // 052-053

  unsigned short hair_style;   // 054-055
  unsigned short weapon;       // 056-057 - see above

  unsigned short baseLv;       // 058-059
  unsigned short skill_pt;     // 060-061 - unallocated skill point

  unsigned short lower_mask;   // 062-063
  unsigned short shield;       // 064-065
  unsigned short helmet;       // 066-067
  unsigned short mid_mask;     // 068-069
  unsigned short hair_color;   // 070-071
  unsigned short cloth_color;  // 072-073

  unsigned char  name[24];     // 074-097
  unsigned char  STR;          // 098
  unsigned char  AGI;          // 099
  unsigned char  VIT;          // 100
  unsigned char  INT;          // 101
  unsigned char  DEX;          // 102
  unsigned char  LUK;          // 103
  unsigned char  index;        // 104
  unsigned char  reserved;     // 105
} msg0x006B;
//--------------------------------------------------------------------
// 0x006C: Bad Character Selection
typedef struct
{
  unsigned short cmd;    // 00-01
} msg0x006C;

//--------------------------------------------------------------------
// 0x006D: Create Character
typedef struct
{
  unsigned short cmd;    // 00-01
} msg0x006D;

//--------------------------------------------------------------------
// 0x006E: Character creation Failed
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned char  code;   // 02   - 0 - name exists, 1 - underaged, 2-255 denied
} msg0x006E;

//--------------------------------------------------------------------
// 0x006F: Delete Character
typedef struct
{
  unsigned short cmd;    // 00-01
} msg0x006F;

//--------------------------------------------------------------------
// 0x0071: Map Server Information
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned long  charID;      // 02-05
  unsigned char  mapname[16]; // 06-21
  struct in_addr ip;          // 22-25
  unsigned short port;        // 26-27
} msg0x0071;

//--------------------------------------------------------------------
// 0x0072: Map Login
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned long  accountID; // 02-05
  unsigned long  charID;    // 06-09
  unsigned long  sessionID; // 10-13
  unsigned long  tick;      // 14-17
  unsigned char  sex;       // 18
} msg0x0072;

//--------------------------------------------------------------------
// 0x0073: Character Initial Position

// coordination format:
//   xxxx xxxx xxyy yyyy yyyy dddd - x, y, direction
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned long  tick;     // 02-05
  unsigned char  coord[3]; // 06-08 - see above
  unsigned short reserved; // 09-10
} msg0x0073;

//--------------------------------------------------------------------
// 0x0075: Unknown
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned char  unknown[9];  // 02-10
} msg0x0075;

//--------------------------------------------------------------------
// 0x0077: Unknown
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned char  unknown[2];  // 02-03
} msg0x0077;

//--------------------------------------------------------------------
// 0x0078: Existing Character
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned long  id;          // 02-05
  unsigned short move_delay;  // 06-07 - delay per block
  unsigned short state;       // 08-09 - see 0x006B
  unsigned short ailments;    // 10-11 - see 0x006B
  unsigned short options;     // 12-13 - see 0x006B

  unsigned short type;        // 14-15

  unsigned short hair_style;  // 16-17  - if value is 20 - it's a pet
  unsigned short weapon;      // 18-19
  unsigned short lower_mask;  // 20-21
  unsigned short shield;      // 22-23
  unsigned short helmet;      // 24-25
  unsigned short mid_mask;    // 26-27
  unsigned short hair_color;  // 28-29
  unsigned short cloth_color; // 30-31

  unsigned short head_dir;    // 32-33

  unsigned long  guild_id;    // 34-37
  unsigned short guild_pos;   // 38-39

  unsigned short manner;      // 40-41
  unsigned short karma;       // 42-43
  unsigned char  stance;      // 44    - 0 - normal, 1 - ready to fight

  unsigned char  sex;         // 45
  unsigned char  coord[3];    // 46-48 - see 0x0073
  unsigned char  unknown[2];  // 49-50
  unsigned char  Sitting;     // 51    - 0 - standing, 1 - lying dead, 2 - sitting
  unsigned short baseLv;      // 52-53
} msg0x0078;

//--------------------------------------------------------------------
// 0x0079: Connected Character
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned long  id;          // 02-05
  unsigned short move_delay;  // 06-07 - delay per block
  unsigned short state;       // 08-09 - see 0x006B
  unsigned short ailments;    // 10-11 - see 0x006B
  unsigned short options;     // 12-13 - see 0x006B

  unsigned short type;        // 14-15

  unsigned short hair_style;  // 16-17 - if value is 20 - it's a pet
  unsigned short weapon;      // 18-19
  unsigned short lower_mask;  // 20-21
  unsigned short shield;      // 22-23
  unsigned short helmet;      // 24-25
  unsigned short mid_mask;    // 26-27
  unsigned short hair_color;  // 28-29
  unsigned short cloth_color; // 30-31

  unsigned short head_dir;    // 32-33

  unsigned long  guild_id;    // 34-37
  unsigned short guild_pos;   // 38-39

  unsigned short manner;      // 40-41
  unsigned short karma;       // 42-43
  unsigned char  stance;      // 44    - 0 - normal, 1 - ready to fight

  unsigned char  sex;         // 45
  unsigned char  coord[3];    // 46-48 - see 0x0073
  unsigned char  unknown[2];  // 49-50
  unsigned short baseLv;      // 51-52
} msg0x0079;

//--------------------------------------------------------------------
// 0x007A: Unknown
typedef struct
{
  unsigned short cmd;        // 00-01
  unsigned char  unknown[2]; // 02-03
} msg0x007A;

//--------------------------------------------------------------------
// 0x007B: Character Move

// coord
//   0         1         2         3         4
//   xxxx xxxx xxyy yyyy yyyy xxxx xxxx xxyy yyyy yyyy
//   |--------- src --------| |--------- dst --------|

// orientation
//   xxxxyyyy

// These packets contain the movement of pet (unnamed) Chon Chon
// it's ID is 00002b7C (7C 2B 00 00)

//21:44:33 Outgoing packet to client: 60 bytes
//0x0000 (0000):7b 00 xx xx xx xx 96 00  00 00 00 00 00 00 f3 03     {.|+.... ........
//0x0010 (0016):14 00 00 00 00 00 55 98  d5 54 00 00 00 00 00 00     ......U. .T......
//0x0020 (0032):00 00 00 00 00 00 00 00  00 00 00 00 00 00 00 00     ........ ........
//0x0030 (0048):00 00 10 8c b1 10 cb 98  00 00 04 00

//21:46:14 Outgoing packet to client: 60 bytes
//0x0000 (0000):7b 00 xx xx xx xx 96 00  00 00 00 00 00 00 f3 03     {.|+.... ........
//0x0010 (0016):14 00 00 00 00 00 9a 23  d7 54 00 00 00 00 00 00     .......# .T......
//0x0020 (0032):00 00 00 00 00 00 00 00  00 00 00 00 00 00 00 00     ........ ........
//0x0030 (0048):00 00 0e cc 20 e4 c1 66  00 00 04 00                 .......f ....

//
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned long  id;          // 02-05
  unsigned short move_delay;  // 06-07 - delay per block
  unsigned short state;       // 08-09 - see 0x006B
  unsigned short ailments;    // 10-11 - see 0x006B
  unsigned short options;     // 12-13 - see 0x006B

  unsigned short type;        // 14-15

  unsigned short hair_style;  // 16-17  - if value is 20 - it's a pet
  unsigned short weapon;      // 18-19
  unsigned short lower_mask;  // 20-21

  unsigned long  tick;        // 22-25

  unsigned short shield;      // 26-27
  unsigned short helmet;      // 28-29
  unsigned short mid_mask;    // 30-31
  unsigned short hair_color;  // 32-33
  unsigned short cloth_color; // 34-35

  unsigned short head_dir;    // 36-37

  unsigned long  guild_id;    // 38-41
  unsigned short guild_pos;   // 42-43

  unsigned short manner;      // 44-45
  unsigned short karma;       // 46-47
  unsigned char  stance;      // 48    - 0 - normal, 1 - ready to fight

  unsigned char  sex;          // 49
  unsigned char  coord[5];     // 50-54 - see above
  unsigned char  orientation;  // 55    - see above
  unsigned char  unknown[2];   // 56-57
  unsigned short baseLv;       // 58-59
} msg0x007B;

//--------------------------------------------------------------------
// 0x007C: NPC respawn
typedef struct
{
  unsigned short cmd;          // 00-01
  unsigned long  id;           // 02-05
  unsigned short move_delay;   // 06-07
  unsigned short state;        // 08-09 - see 0x006B
  unsigned short ailments;     // 10-11 - see 0x006B
  unsigned short options;      // 12-13 - see 0x006B
  unsigned short lower_mask;   // 14-15

  unsigned long  tick;         // 16-19

  unsigned short type;         // 20-21

  unsigned char  unknown2[13]; // 22-34
  unsigned char  sex;          // 35
  unsigned char  coord[3];     // 36-38 // see 0x0073
  unsigned char  unknown3[2];  // 39-40
} msg0x007C;

//--------------------------------------------------------------------
// 0x007D: Confirm Map is Loaded
typedef struct
{
  unsigned short cmd;    // 00-01
} msg0x007D;

//--------------------------------------------------------------------
// 0x007E: Time Sycnhronisation
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned long  time; // 02-05
} msg0x007E;

//--------------------------------------------------------------------
// 0x007F: Time Synchronisation
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned long  time; // 02-05
} msg0x007F;

//--------------------------------------------------------------------
// 0x0080: Character Died
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned long  ID;   // 02-05
  unsigned char  type; // 06
} msg0x0080;

//--------------------------------------------------------------------
// 0x0081: Disconnected from Server
typedef struct
{
  unsigned short cmd;  // 00-01;
  unsigned char  code; // 02
} msg0x0081;

//--------------------------------------------------------------------
// 0x0085: Moving
//
// coordination format:
//   xxxx xxxx xxyy yyyy yyyy dddd - x, y, direction
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned char  coord[3]; // 02-04
} msg0x0085;

//--------------------------------------------------------------------
// 0x0087: Character Moved
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned long  tick;        // 02-05
  unsigned char  coord[5];    // 07-10
  unsigned char  orientation; // 11
} msg0x0087;

//--------------------------------------------------------------------
// 0x0088: New character position
typedef struct
{
  unsigned short cmd; // 00-01
  unsigned long  id;  // 02-05
  unsigned short x;   // 06-07
  unsigned short y;   // 08-09
} msg0x0088;

//--------------------------------------------------------------------
// 0x0089: Action
typedef struct
{
// action code
//  00 - attack once
//  02 - sitdown
//  03 - standup
//  07 - lock target

  unsigned short cmd;    // 00-01
  unsigned long  target; // 02-05
  unsigned char  action; // 06 07
} msg0x0089;

//--------------------------------------------------------------------
// 0x008A: Action to Character

// action type
//   00 - single attack
//   01 - pickup item
//   02 - sit down
//   03 - stand up
//   04 - attack to endured target
//   08 - multiple attack
//   0a - critical attack
//   0b - perfect dodge

typedef struct
{
  unsigned short cmd;          // 00-01
  unsigned long  srcID;        // 02-05
  unsigned long  dstID;        // 06-09
  unsigned long  tick;         // 10-13
  unsigned long  src_speed;    // 14-17
  unsigned long  dst_speed;    // 18-21
  unsigned short damage;       // 22-23
  unsigned short attack_count; // 24-25 - exclude lefthand damage
  unsigned char  actiontype;   // 26    - see above
  unsigned short lefthand_dmg; // 27-28
} msg0x008A;

//--------------------------------------------------------------------
// 0x008B: Action to Character
typedef struct
{
  unsigned short cmd;          // 00-01
  unsigned long  srcID;        // 02-05
  unsigned long  dstID;        // 06-09
  unsigned long  tick;         // 10-13
  unsigned long  src_speed;    // 14-17
  unsigned long  dst_speed;    // 18-21
  unsigned short damage;       // 22-23
  unsigned short attack_count; // 24-25 - exclude lefthand damage
  unsigned char  actiontype;   // 26    - see 0x008A
  unsigned short lefthand_dmg; // 27-28
} msg0x008B;

//--------------------------------------------------------------------
// 0x008C: Chat message
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short msglen; // 02-03
  unsigned char  msg[];  // 04-
} msg0x008C;

//--------------------------------------------------------------------
// 0x008D: Chat message
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short msglen; // 02-03
  unsigned long  ID;     // 04-07
  unsigned char  msg[];  // 08-
} msg0x008D;

//--------------------------------------------------------------------
// 0x008E: Chat message
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short msglen; // 02-03
  unsigned char  msg[];  // 04-
} msg0x008E;

//--------------------------------------------------------------------
// 0x008F: Unknown
typedef struct
{
  unsigned short cmd;        // 00-01
  unsigned char  unknown[2]; // 02-03
} msg0x008F;

//--------------------------------------------------------------------
// 0x0090: Talk to NPC
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned long  npcID;  // 02-05
  unsigned char  flag;   // 06
} msg0x0090;

//--------------------------------------------------------------------
// 0x0091: Map Changed
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned char  mapname[16]; // 02-17
  unsigned short x;     // 18-19
  unsigned short y;     // 20-21
} msg0x0091;

//--------------------------------------------------------------------
// 0x0092: Map Server Changed
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned char  mapname[16]; // 02-17
  unsigned short x;           // 18-19
  unsigned short y;           // 20-21
  struct in_addr ip;          // 22-25
  unsigned short port;        // 26-27
} msg0x0092;

//--------------------------------------------------------------------
// 0x0093: Unknown
typedef struct
{
  unsigned short cmd;      // 00-01
} msg0x0093;

//--------------------------------------------------------------------
// 0x0094: Request for Player Information
typedef struct
{
  unsigned short cmd;        // 00-01
  unsigned long  accountID;  // 02-05
} msg0x0094;

//--------------------------------------------------------------------
// 0x0095: Character's name
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned long  ID;       // 02-05
  unsigned char  name[24]; // 06-29
} msg0x0095;


//--------------------------------------------------------------------
// 0x0096: Send private message
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short msglen;   // 02-03
  unsigned char  name[24]; // 04-27
  unsigned char  msg[];    // 28-
} msg0x0096;

//--------------------------------------------------------------------
// 0x0097: Incoming Whisper
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short msglen;   // 02-03
  unsigned char  from[24]; // 04-27
  unsigned char  msg[];    // 28-
} msg0x0097;

//--------------------------------------------------------------------
// 0x0098: Whisper Result
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned char  result; // 02
} msg0x0098;

//--------------------------------------------------------------------
// 0x009A: Chat (9A)
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short msglen; // 02-03
  unsigned char  msg[];  // 04-
} msg0x009A;

//--------------------------------------------------------------------
// 0x009B: Look At (outgoing)
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned short head;    // 02-03 - 0 forward, 1 left, 2 right
  unsigned char  body;    // 04    - 0 north, 1 - north-west, 2 - west, 3 - south-west, 4 - south, 5 - south-east, 6 - east, 7 - north-west 
} msg0x009B;

//--------------------------------------------------------------------
// 0x009C: Look At (incoming)
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned long  ID;      // 02-05
  unsigned short head;    // 06-07 - 0 forward, 1 left, 2 right
  unsigned char  body;    // 08    - see 0x009B
} msg0x009C;

//--------------------------------------------------------------------
// 0x009D: Item Exists
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned long  ID;       // 02-05
  unsigned short type;     // 06-07
  unsigned char  nonequip; // 08
  unsigned short x;        // 09-10
  unsigned short y;        // 11-12
  unsigned short amount;   // 13-14
  unsigned short unknown;  // 15-16 - sub x, sub y ???
} msg0x009D;

//--------------------------------------------------------------------
// 0x009E: Item Appears
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned long  ID;       // 02-05
  unsigned short type;     // 06-07
  unsigned char  nonequip; // 08
  unsigned short x;        // 09-10
  unsigned short y;        // 11-12
  unsigned short unknown;  // 13-14
  unsigned short amount;   // 15-16 - sub x, sub y ???
} msg0x009E;

//--------------------------------------------------------------------
// 0x009F: Pick Item
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned long  itemID; // 02-05
} msg0x009F;

//--------------------------------------------------------------------
// 0x00A0: Item Picked Up
typedef struct
{
  unsigned short cmd;        // 00-01

  union
  {
    struct
    {
      unsigned short index;  // 02-03
      unsigned short amount; // 04-05
    } inventory;

    unsigned long id;        // 02-05

  } info;

  unsigned short type;       // 06-07
  unsigned char  identified; // 08
  unsigned char  element;    // 09 - 00 - Neutral, 01 - Water, 02 - Earth, 03 - Fire, 04 Wind
  unsigned char  refine;     // 10

//  unsigned short slot[4];    // 11-18
  union
  {
    unsigned short slot[4];

    struct
    {
      unsigned short flag;     // always = 0x00FF
      unsigned char  element;  // 1 - Water, 2 - Earth, 3 - Fire, 4 - Wind
      unsigned char  strength; // 5 - very strong, 10 - very very strong
      unsigned long  bsid;     // id of BS who build this weapon
    } smitten;

  } attributes;               // 11-18

  unsigned short equiptype;  // 19-20

  unsigned char  category;   // 21
  unsigned char  result;     // 22

  // result:
  //   00 - success
  //   01 - was picked by others
  //   02 - Overweight
  //   03 - success - add to existing item in inventory
  //   04 - ??
  //   05 - cannot have more than 30,000 pieces in inventory
  //   06 - anti-looted
} msg0x00A0;

//--------------------------------------------------------------------
// 0x00A1: Item Disappeared
typedef struct
{
  unsigned short cmd; // 00-01
  unsigned long  ID;  // 02-05
} msg0x00A1;

//--------------------------------------------------------------------
// 0x00A2: Drop Item
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short index;  // 02-03
  unsigned long  amount; // 04-07
} msg0x00A2;

//--------------------------------------------------------------------
// 0x00A3: Item in the Inventory
typedef struct
{
  unsigned short index;      // 00-01
  unsigned short type;       // 02-03
  unsigned char  category;   // 04
  unsigned char  identified; // 05
  unsigned short amount;     // 06-07
  unsigned short unknown;    // 08-09
} msg0x00A3;

//--------------------------------------------------------------------
// 0x00A4: Equipment in the Inventory

// 00 01 02 03 04 05 06 07 08 09 10 11 12 13 14 15 16 17 18 19
// -----------------------------------------------------------
// xx 00 82 05 09 01 00 00 00 00 00 05 ff 00 01 0a xx xx xx xx +5 Very Very Strong <name>'s Ice Lance
// xx 00 e0 05 05 01 00 00 00 00 00 08 ff 00 03 05 xx xx xx xx +8 Very Strong <name>'s Fire Mace
// xx 00 82 05 09 01 00 00 00 00 00 05 ff 00 03 0a xx xx xx xx +5 Very Very Strong <name>'s Fire Lance
// xx 00 82 05 09 01 00 00 00 00 00 05 ff 00 04 0a xx xx xx xx +5 Very Very Strong <name>'s Wind Lance

typedef struct
{
  unsigned short index;      // 00-01
  unsigned short type;       // 02-03
  unsigned char  category;   // 04
  unsigned char  identified; // 05
  unsigned short equiptype;  // 06-07
  unsigned short equipping;  // 08-09
  unsigned char  element;    // 10
  unsigned char  refine;     // 11
//  unsigned short slot[4];    // 12-19
  union
  {
    unsigned short slot[4];

    struct
    {
      unsigned short flag;     // always = 0x00FF
      unsigned char  element;  // 1 - Water, 2 - Earth, 3 - Fire, 4 - Wind
      unsigned char  strength; // 5 - very strong, 10 - very very strong
      unsigned long  bsid;     // id of BS who build this weapon
    } smitten;

  } attributes;               // 12-19
} msg0x00A4;
//--------------------------------------------------------------------
// 0x00A5: Item in Storage
typedef struct
{
  unsigned short index;      // 00-01
  unsigned short type;       // 02-03
  unsigned char  category;   // 04
  unsigned char  identified; // 05 << always 1?
  unsigned long  amount;     // 06-09
} msg0x00A5;
//--------------------------------------------------------------------
// 0x00A6: Equipment in Storage
typedef struct
{
  unsigned short index;       // 00-01
  unsigned short type;        // 02-03
  unsigned char  category;    // 04
  unsigned char  identified;  // 05
  unsigned short equiptype;   // 06-07
  unsigned short equipping;   // 08-09
  unsigned char  element;     // 10
  unsigned char  refine;      // 11

  union
  {
    unsigned short slot[4];

    struct
    {
      unsigned short flag;     // always = 0x00FF
      unsigned char  element;  // 1 - Water, 2 - Earth, 3 - Fire, 4 - Wind
      unsigned char  strength; // 5 - very strong, 10 - very very strong
      unsigned long  bsid;     // id of BS who build this weapon
    } smitten;

  } attributes;               // 12-19
} msg0x00A6;

//--------------------------------------------------------------------
// 0x00A7: Use item
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short index;  // 02-03
  unsigned long  accID;  // 04-07
} msg0x00A7;

//--------------------------------------------------------------------
// 0x00A8: Item used
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned short index;     // 02-03
  unsigned short remaining; // 04-05
  unsigned char  amount;    // 06
} msg0x00A8;

//--------------------------------------------------------------------
// 0x00A9: Equip equipment
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned short index;     // 02-03
  unsigned short equiptype; // 04-05
} msg0x00A9;

//--------------------------------------------------------------------
// 0x00AA: Equipment equipped
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned short index;     // 02-03
  unsigned short equipping; // 04-05
  unsigned char  result;    // 06 : 1 = success
} msg0x00AA;

//--------------------------------------------------------------------
// 0x00AB: Unequip equipment
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short index;  // 02-03
} msg0x00AB;

//--------------------------------------------------------------------
// 0x00AC: Equipment unequipped
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned short index;     // 02-03
  unsigned short equipping; // 04-05
  unsigned char  result;    // 06 : 1 = successfully unequipped
} msg0x00AC;

//--------------------------------------------------------------------
// 0x00AF: Item Dropped
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned short index;   // 02-03
  unsigned short amount;  // 04-05
} msg0x00AF;

//--------------------------------------------------------------------
// 0x00B0: Got Player Status
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned short type;    // 02-03
  unsigned long  value;   // 04-07
} msg0x00B0;

//--------------------------------------------------------------------
// 0x00B1: Status Update
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned short type;  // 02-03
  signed   long  value; // 04-07
} msg0x00B1;

//--------------------------------------------------------------------
// 0x00B2: Respawn
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned char  type; // 02   - 0 - respawn, 1 - reselect char
} msg0x00B2;

//--------------------------------------------------------------------
// 0x00B3: Reselect Character
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned char  unknown1; // 02
  unsigned long  unknown2; // 03-06
} msg0x00B3;

//--------------------------------------------------------------------
// 0x00B4: NPC Talk
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short msglen; // 02-03
  unsigned long  ID;  // 04-07
  unsigned char  msg[];  // 08-
} msg0x00B4;

//--------------------------------------------------------------------
// 0x00B5: NPC wait
typedef struct
{
  unsigned short cmd; // 00-01
  unsigned long  ID;  // 02-05
} msg0x00B5;

//--------------------------------------------------------------------
// 0x00B6: NPC Stop Talk
typedef struct
{
  unsigned short cmd; // 00-01
  unsigned long  ID;  // 02-05
} msg0x00B6;

//--------------------------------------------------------------------
// 0x00B7: NPC wait for choice
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short msglen; // 02-03
  unsigned long  ID;  // 04-07
  unsigned char  msg[];  // 08-
} msg0x00B7;

//--------------------------------------------------------------------
// 0x00B8: Response to NPC
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned long  npcID;    // 02-05
  unsigned char  response; // 06
} msg0x00B8;

//--------------------------------------------------------------------
// 0x00B9: Continue talk to NPC
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned long  npcID;  // 02-05
} msg0x00B9;

//--------------------------------------------------------------------
// 0x00BB: Add Status Point

// type
//   0x0000 - STR
//   0x000E - AGI
//   0x000F - VIT
//   0x0010 - INT
//   0x0011 - DEX
//   0x0012 - INT
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short type;   // 02-03
  unsigned char  amount; // 04
} msg0x00BB;

//--------------------------------------------------------------------
// 0x00BC: Status Add Result
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short type;   // 02-03 - see 0x00BB
  unsigned char  result; // 04
  unsigned char  value;  // 05 
} msg0x00BC;

//--------------------------------------------------------------------
// 0x00BD: Calculated Status
typedef struct
{
  unsigned short cmd;          // 00-01
  unsigned short point_avail;  // 02-03
  unsigned char  STR;          // 04
  unsigned char  STR_required; // 05
  unsigned char  AGI;          // 06
  unsigned char  AGI_required; // 07
  unsigned char  VIT;          // 08
  unsigned char  VIT_required; // 09
  unsigned char  INT;          // 10
  unsigned char  INT_required; // 11
  unsigned char  DEX;          // 12
  unsigned char  DEX_required; // 13
  unsigned char  LUK;          // 14
  unsigned char  LUK_required; // 15
  unsigned short ATTK;         // 16-17
  unsigned short ATTK_bonus;   // 18-19
  unsigned short MATK_min;     // 20-21
  unsigned short MATK_max;     // 22-23
  unsigned short DEF;          // 24-25
  unsigned short DEF_bonus;    // 26-27
  unsigned short MDEF;         // 28-29
  unsigned short MDEF_bonus;   // 30-31
  unsigned short Hit;          // 32-33
  unsigned short Flee;         // 34-35
  unsigned short Flee_bonus;   // 36-37
  unsigned short Critical;     // 38-39
  unsigned char  unknown[4];   // 40-43
} msg0x00BD;

//--------------------------------------------------------------------
// 0x00BE: Status Point Required
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned short type;  // 02-03
  unsigned char  value; // 04
} msg0x00BE;

//--------------------------------------------------------------------
// 0x00BF: Send emoticon
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned char  emoticon; // 02
} msg0x00BF;

//--------------------------------------------------------------------
// 0x00C0: Got Emoticon
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned long  ID;   // 02-05
  unsigned char  emotion; // 06
} msg0x00C0;

//--------------------------------------------------------------------
// 0x00C1: Request number of online users ?????
typedef struct
{
  unsigned short cmd;   // 00-01
} msg0x00C1;

//--------------------------------------------------------------------
// 0x00C2: Number of online users
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned long  users; // 02-05
} msg0x00C2;

//--------------------------------------------------------------------
// 0x00C3: Change Job
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned long  AccID;    // 02-05
  unsigned char  JobID;    // 06
  unsigned char  NewJobID; // 07
} msg0x00C3;

//--------------------------------------------------------------------
// 0x00C4: Request NPC Trade
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned long  npcID;   // 02-05
} msg0x00C4;

//--------------------------------------------------------------------
// 0x00C5: List item in store / List sell-able item
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned long  npcID; // 02-05
  unsigned char  mode;  // 06: 0 = to buy, 1 = to sell
} msg0x00C5;

//--------------------------------------------------------------------
// 0x00C6: List Sell items
typedef struct
{
  unsigned long  price;    // 00-03
  unsigned long  dc_price; // 04-07
  unsigned char  category; // 08
  unsigned short type;     // 09-10
} msg0x00C6;

//--------------------------------------------------------------------
// 0x00C7: Ready for Selling
typedef struct
{
  unsigned short index;    // 00-01
  unsigned long  price;    // 02-03
  unsigned long  oc_price; // 04-07
} msg0x00C7;

//--------------------------------------------------------------------
// 0x00C8: Buy
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short msglen; // 02-03
  struct
  {
    unsigned short amount;    // 04-05
    unsigned short item_type; // 06-07 ...
  } item[];
  //...
} msg0x00C8;

//--------------------------------------------------------------------
// 0x00C9: Sell
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short msglen; // 02-03
  struct
  {
    unsigned short index;  // 04-05
    unsigned short amount; // 06-07 ...
  } item[];
} msg0x00C9;

//--------------------------------------------------------------------
// 0x00CA: Done buying?
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned char  flag; // 02
} msg0x00CA;

//--------------------------------------------------------------------
// 0x00CB: Done Selling?
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned char  flag; // 02
} msg0x00CB;

//--------------------------------------------------------------------
// 0x00CF: Ignore
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned char  name[24]; // 02-25
  unsigned char  flag;     // 26
} msg0x00CF;

//--------------------------------------------------------------------
// 0x00D0: Ignore all
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned char  flag;  // 02
} msg0x00D0;

//--------------------------------------------------------------------
// 0x00D1: Ignore player result
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned char  type;  // 02
  unsigned char  error; // 03
} msg0x00D1;

//--------------------------------------------------------------------
// 0x00D2: Ignore all result
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned char  type;  // 02
  unsigned char  error; // 03
} msg0x00D2;

//--------------------------------------------------------------------
// 0x00D3: Get ignore list ?????
//   type: Bidirection
typedef struct
{
  unsigned short cmd;   // 00-01
} msg0x00D3;

//--------------------------------------------------------------------
// 0x00D5: Create Chat Room
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned short msglen;    // 02-03
  unsigned short limit;     // 04-05
  unsigned char  isPublic;  // 06
  unsigned char  passwd[8]; // 07-14
  unsigned char  title[];   // 15-
} msg0x00D5;

//--------------------------------------------------------------------
// 0x00D6: Chat room created
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned char  result; // 03
} msg0x00D6;

//--------------------------------------------------------------------
// 0x00D7: Chat room information
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short msglen;   // 02-03
  unsigned long  ownerID;  // 04-07
  unsigned long  roomID;   // 08-11
  unsigned short limit;    // 12-13
  unsigned short users;    // 14-15
  unsigned char  isPublic; // 16
  unsigned char  title[];  // 17-
} msg0x00D7;

//--------------------------------------------------------------------
// 0x00D8: Remove Chat Room
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned long  roomID; // 02-05
} msg0x00D8;

//--------------------------------------------------------------------
// 0x00D9: Join Chat Room
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned long  ID;        // 02-05
  unsigned char  passwd[8]; // 06-13
} msg0x00D9;

//--------------------------------------------------------------------
// 0x00DA: Cannot Join Chat Room
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned char  result; // 02
} msg0x00DA;

//--------------------------------------------------------------------
// 0x00DB: Join Chat Room
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short msglen; // 02-03
  unsigned long  roomID; // 04-07
} msg0x00DB;

typedef struct
{
  unsigned char type;       // 00
  unsigned char unknown[3]; // 01-03
  unsigned char name[24];   // 04-27
} msg0x00DBex;

//--------------------------------------------------------------------
// 0x00DC: Other player join chat room
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short users;    // 02-03
  unsigned char  name[24]; // 04-27
} msg0x00DC;

//--------------------------------------------------------------------
// 0x00DD: Other player leave chat room
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned char  unknown1[2]; // 02-03
  unsigned char  name[24];    // 04-27
  unsigned char  unknown2;    // 28
} msg0x00DD;

//--------------------------------------------------------------------
// 0x00DE: Change chat room properties
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned short msglen;    // 02-03
  unsigned short limit;     // 04-05
  unsigned char  isPublic;  // 06
  unsigned char  passwd[8]; // 07-14
  unsigned char  title[];   // 15-
} msg0x00DE;

//--------------------------------------------------------------------
// 0x00DF: Change chat room properties
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short msglen;   // 02-03
  unsigned long  ownerID;  // 04-07
  unsigned long  roomID;   // 08-11
  unsigned short limit;    // 12-13
  unsigned short users;    // 14-15
  unsigned char  isPublic; // 16
  unsigned char  title[];  // 17-
} msg0x00DF;

//--------------------------------------------------------------------
// 0x00E0: Bestow Chat Room
typedef struct
{
  unsigned short cmd;        // 00-01
  unsigned char  unknown[4]; // 02-05
  unsigned char  name[24];   // 06-29
} msg0x00E0;

//--------------------------------------------------------------------
// 0x00E1: Chat room owner
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned char  type;        // 02
  unsigned char  unknown1[3]; // 03-05
  unsigned char  name[24];    // 06-29
} msg0x00E1;

//--------------------------------------------------------------------
// 0x00E2: Kick player from chat room
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned char  name[24]; // 02-25
} msg0x00E2;

//--------------------------------------------------------------------
// 0x00E3: Leave current chat room
typedef struct
{
  unsigned short cmd;   // 00-01
} msg0x00E3;

//--------------------------------------------------------------------
// 0x00E4: Initiate Deal
typedef struct
{
  unsigned short cmd; // 00-01
  unsigned long  ID;  // 02-05
} msg0x00E4;

//-------------------------------------------------------------------
// 0x00E5: Incoming deal
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned char  name[24]; // 02-25
} msg0x00E5;

//--------------------------------------------------------------------
// 0x00E6: Accept/Cancel the deal
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned char  decision; // 02: 03 = accept, 04 = cancel
} msg0x00E6;

//--------------------------------------------------------------------
// 0x00E7: Engage the deal
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned char  type; // 02
} msg0x00E7;

//--------------------------------------------------------------------
// 0x00E8: Add item to the deal
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short index;  // 02-03
  unsigned long  amount; // 04-07
} msg0x00E8;

//--------------------------------------------------------------------
// 0x00E9: Item added to the deal
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned long  amount;      // 02-05
  unsigned short type;        // 06-07
  unsigned char  unknown[11]; // 08-18
} msg0x00E9;

//--------------------------------------------------------------------
// 0x00EA: Submit the deal
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short index;    // 02-03
  unsigned char  unknown;  // 04
} msg0x00EA;

//--------------------------------------------------------------------
// 0x00EB: Finalise the deal
typedef struct
{
  unsigned short cmd;   // 00-01
} msg0x00EB;

//--------------------------------------------------------------------
// 0x00EC: Deal is finalised
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned char  type; // 02
} msg0x00EC;

//--------------------------------------------------------------------
// 0x00ED: Cancel current deal
typedef struct
{
  unsigned short cmd;   // 00-01
} msg0x00ED;

//--------------------------------------------------------------------
// 0x00EE: Deal canceled
typedef struct
{
  unsigned short cmd;  // 00-01
} msg0x00EE;

//--------------------------------------------------------------------
// 0x00EF: Trade the deal
typedef struct
{
  unsigned short cmd;   // 00-01
} msg0x00EF;

//--------------------------------------------------------------------
// 0x00F0: Deal completed
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned char unknown; // 02
} msg0x00F0;

//--------------------------------------------------------------------
// 0x00F2: Got list of item in storage
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned short items;     // 02-03
  unsigned short items_max; // 04-05
} msg0x00F2;

//--------------------------------------------------------------------
// 0x00F3: Add item to storage
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short index;  // 02-03
  unsigned long  amount; // 04-07
} msg0x00F3;

//--------------------------------------------------------------------
// 0x00F4: Item added to storage
typedef struct
{
  unsigned short cmd;        // 00-01
  unsigned short index;      // 02-03
  unsigned long  amount;     // 04-07
  unsigned short type;       // 08-09

  unsigned char  identified; // 10
  unsigned char  element;    // 11
  unsigned char  refine;     // 12
//  unsigned short slot[4];    // 13-20
  union
  {
    unsigned short slot[4];

    struct
    {
      unsigned short flag;     // always = 0x00FF
      unsigned char  element;  // 1 - Water, 2 - Earth, 3 - Fire, 4 - Wind
      unsigned char  strength; // 5 - very strong, 10 - very very strong
      unsigned long  bsid;     // id of BS who build this weapon
    } smitten;

  } attributes;               // 13-20
} msg0x00F4;

//--------------------------------------------------------------------
// 0x00F5: Get item from storage
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short index;  // 02-03
  unsigned long  amount; // 04-07
} msg0x00F5;

//--------------------------------------------------------------------
// 0x00F6: Item removed from storage (response to 0x00F5)
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short index;  // 02-03
  unsigned long  amount; // 04-07
} msg0x00F6;

//--------------------------------------------------------------------
// 0x00F7: Close storage
typedef struct
{
  unsigned short cmd;   // 00-01
} msg0x00F7;

//--------------------------------------------------------------------
// 0x00F8: Storage closed
typedef struct
{
  unsigned short cmd;  // 00-01
} msg0x00F8;

//--------------------------------------------------------------------
// 0x00F9: Create (organize) a party
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned char  name[24]; // 02-25
} msg0x00F9;

//--------------------------------------------------------------------
// 0x00FA: Result of party creation
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned char  result; // 02
} msg0x00FA;

//--------------------------------------------------------------------
// 0x00FB: Party member's information
typedef struct
{
  unsigned short cmd;           // 00-01
  unsigned short msglen;        // 02-03
  unsigned char  partyname[24]; // 04-27
} msg0x00FB;

typedef struct
{
  unsigned long ID;       // 00-03
  unsigned char name[24]; // 04-27
  unsigned char map[16];  // 28-43
  unsigned char order;    // 44
  unsigned char offline;  // 45
} msg0x00FBex;

//--------------------------------------------------------------------
// 0x00FC: Ask player to join party
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned long  playerID; // 02-05
} msg0x00FC;

//--------------------------------------------------------------------
// 0x00FD: Result of asking to join party
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned char  name[24]; // 02-25
  unsigned char  result;   // 26
} msg0x00FD;

//--------------------------------------------------------------------
// 0x00FE: Be invited to join a party
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned long  partyID;  // 02-05
  unsigned char  name[24]; // 06-29
} msg0x00FE;

//--------------------------------------------------------------------
// 0x00FF: Join a party
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned long  partyID;  // 02-05
  unsigned long  flag;     // 06-09
} msg0x00FF;

//--------------------------------------------------------------------
// 0x0100: Leave party
typedef struct
{
  unsigned short cmd;   // 00-01
} msg0x0100;

//--------------------------------------------------------------------
// 0x0101: Got party exp share mode
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short expmode;  // 02-03
  unsigned short reserved; // 04-05
} msg0x0101;

//--------------------------------------------------------------------
// 0x0102: Set Exp share mode
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short expmode;  // 02-03
  unsigned short reserved; // 04-05
} msg0x0102;

//--------------------------------------------------------------------
// 0x0103: Kick party member out
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned long  partyID;  // 02-05
  unsigned char  name[24]; // 06-29
} msg0x0103;

//--------------------------------------------------------------------
// 0x0104: New member join party
typedef struct
{
  unsigned short cmd;           // 00-01
  unsigned long  ID;            // 02-05
  unsigned long  unknown1;      // 06-09
  unsigned short x;             // 10-11
  unsigned short y;             // 12-13
  unsigned char  isNotOnline;   // 14
  unsigned char  partyName[24]; // 15-38
  unsigned char  name[24];      // 39-62
  unsigned char  mapName[16];   // 63-78
} msg0x0104;

//--------------------------------------------------------------------
// 0x0105: Member leave party
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned long  ID;       // 02-05
  unsigned char  name[24]; // 06-29
  unsigned char  result;   // 30
} msg0x0105;

//--------------------------------------------------------------------
// 0x0106: Got party member max HP
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned long  ID;    // 02-05
  unsigned short HP;    // 06-07
  unsigned short MaxHP; // 08-09
} msg0x0106;

//--------------------------------------------------------------------
// 0x0107: Member party moved
typedef struct
{
  unsigned short cmd; // 00-01
  unsigned long  ID;  // 02-05
  unsigned short x;   // 06-07
  unsigned short y;   // 08-09
} msg0x0107;

//--------------------------------------------------------------------
// 0x0108: Party Chat
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short msglen; // 02-03
  unsigned char  msg[];  // 04-
} msg0x0108;

//--------------------------------------------------------------------
// 0x0109: Received party chat
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned short msglen;  // 02-03
  unsigned long  from;    // 04-07
  unsigned char  msg[];   // 08-
} msg0x0109;

//--------------------------------------------------------------------
// 0x010B: MVP Experience
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned long  exp_val; // 02-05
} msg0x010B;

//--------------------------------------------------------------------
// 0x010E: Got skill level
typedef struct
{
  unsigned short cmd;        // 00-01
  unsigned short skillID;    // 02-03
  unsigned short level;      // 04-05
  unsigned char  unknown[5]; // 06-10
} msg0x010E;

//--------------------------------------------------------------------
// 0x010F: Skill Information
typedef struct
{
  unsigned short skillID;       // 00-01
  unsigned short skilltype;     // 02-03
  unsigned short unknown1;      // 04-05
  unsigned short level;         // 06-07
  unsigned short sp;            // 08-09
  unsigned short range;         // 10-11
  unsigned char  skillName[24]; // 12-35
  unsigned char  unknown2;      // 36
} msg0x010F;

//--------------------------------------------------------------------
// 0x0110: Failed to use skill
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short skill_id; // 02-03
  unsigned short level;    // 04-05
  unsigned short unknown;  // 06-07
  unsigned char  result;   // 08    - if 0 - skill is failed
  unsigned char  reason;   // 09

  // reason:
  //   00 - skill dependant
  //   01 - insufficient HP
  //   02 - insufficient SP
  //   03 - no memo (warp)
  //   04 - skill delay
  //   05 - zeny
  //   06 - cannot use skill with current weapon
  //   07 - need red gemstone
  //   08 - need blue gemstone
  //   09 - overweight
  //   0a - casted but failed
  
} msg0x0110;

//--------------------------------------------------------------------
// 0x0111: Skill Info (additional skills)
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned short skill_id;  // 02-03
  unsigned short skilltype; // 04-05
  unsigned short unknown1;  // 06-07
  unsigned short level;     // 08-09
  unsigned short SP;        // 10-11
  unsigned short range;     // 12-13
  unsigned char  name[24];  // 14-37
  unsigned char  unknown2;  // 38
} msg0x0111;

//--------------------------------------------------------------------
// 0x0112: Add Skill Point
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned short skillID; // 02-03
} msg0x0112;

//--------------------------------------------------------------------
// 0x0113: Use Skill
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short level;    // 02-03
  unsigned short skill_id; // 04-05
  unsigned long  dst_id;   // 06-09
} msg0x0113;

//--------------------------------------------------------------------
// 0x0114: Use Skill Approach Action

// type
//   04 - firewall?
//   06 - single hit
//   08 - multiple hit
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned short skill_id;  // 02-03
  unsigned long  src_id;    // 04-07
  unsigned long  dst_id;    // 08-11

  unsigned long  tick;      // 12-15
  unsigned long  src_speed; // 16-19
  unsigned long  dst_speed; // 20-23

  unsigned short damage;    // 24-25
  unsigned short level;     // 26-27
  unsigned short hitnum;    // 28-29
  unsigned char  type;      // 30    - see 0x0114
} msg0x0114;

//--------------------------------------------------------------------
// 0x0115: Skill Effect
// Note: Original information - length = 16
//       Additional info.     - length = 23

// type
//   05 - damage blow-up
//   06 - explosion
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned short skill_id;  // 02-03
  unsigned long  source_id; // 04-07
  unsigned long  dest_id;   // 08-11
  unsigned short x;         // 12-13
  unsigned short y;         // 14-15
  unsigned short damage;    // 16-17
  unsigned short level;     // 18-19
  unsigned short hitnum;    // 20-21
  unsigned char  type;      // 22
} msg0x0115;

//--------------------------------------------------------------------
// 0x0116: Use skill at specific location
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short level;    // 02-03
  unsigned short skill_id; // 04-05
  unsigned short x;        // 06-07
  unsigned short y;        // 08-09
} msg0x0116;

//--------------------------------------------------------------------
// 0x0117: Skill used at specific location
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short skill_id; // 02-03
  unsigned long  src_id;   // 04-07
  unsigned short level;    // 08-09
  unsigned short x;        // 10-11
  unsigned short y;        // 12-13
  unsigned long  tick;     // 14-17
} msg0x0117;

//--------------------------------------------------------------------
// 0x0118: Stop Attack
typedef struct
{
  unsigned short cmd;    // 00-01
} msg0x0118;

//--------------------------------------------------------------------
// 0x0119: StatusChanged

// option1 ??
//

// ailments ??
//   0000 0000 0000 0001 0000 0000 0000 0000 - Poison
//   0000 0000 0000 0010 0000 0000 0000 0000 - Curse
//   0000 0000 0000 0100 0000 0000 0000 0000 - Silence
//   0000 0000 0000 1000 0000 0000 0000 0000 - Confusion
//   0000 0000 0001 0000 0000 0000 0000 0000 - blind

// option3 ??
//   0000 0000 0000 0001 - Sight / Ruwach
//   0000 0000 0000 0010 - Hide
//   0000 0000 0000 0100 - Cloak
//   0000 0000 0000 1000 - Cart (Lv 1-40)
//   0000 0000 0001 0000 - Falcon
//   0000 0000 0010 0000 - Pecopeco
//   0000 0000 0100 0000 - Disappear
//   0000 0000 1000 0000 - Cart (Lv 41-65)
//   0000 0001 0000 0000 - Cart (Lv 66-80)
//   0000 0010 0000 0000 - Cart (Lv 81-90)
//   0000 0100 0000 0000 - Cart (Lv 91-99)
//   0000 1000 0000 0000 - Hideous Mark

typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned long  id;       // 02-05
  unsigned long  ailments; // 06-09
  unsigned short options;  // 10-11
  unsigned char  unknown;  // 12
} msg0x0119;

//--------------------------------------------------------------------
// 0x011A: Cast skill that gain status
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short skill_id; // 02-03
  unsigned short amount;   // 04-05 - amount of HP cured ( if healing )
  unsigned long  dst_id;   // 06-09
  unsigned long  src_id;   // 10-13
  unsigned char  result;   // 14
} msg0x011A;

//--------------------------------------------------------------------
// 0x011B: Teleport - Return to saved map
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned short skill_id;    // 02-03 - 0x001A for teleport
  unsigned char  mapname[16]; // 04-19 - if mapname = "cancel" - stop using skill
} msg0x011B;

//--------------------------------------------------------------------
// 0x011C: Place for teleport or portal warp
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short skill_id; // 02-03
  unsigned char  map1[16]; // 04-19
  unsigned char  map2[16]; // 20-35
  unsigned char  map3[16]; // 36-51
  unsigned char  map4[16]; // 52-67
} msg0x011C;

//--------------------------------------------------------------------
// 0x011D: Memorise warp portal location
typedef struct
{
  unsigned short cmd;    // 00-01
} msg0x011D;

//--------------------------------------------------------------------
// 0x011E: Warp portal location memorised
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned char  isFailed; // 02
} msg0x011E;

//--------------------------------------------------------------------
// 0x011F: Area affected by skill

// type (parts of)
//  0x7E - ??
//  0x7F - Firewall
//  0x80 - Warp Portal - invoking
//  0x81 - Warp Portal - before invoke
//  0x8c - talkie box  - invoking
//  0x91 - Ankle Snare
//  0x93 - Land mine
//  0x99 - Talkie Box - before invoke
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned long  src_id;   // 02-05
  unsigned long  skill_id; // 06-09
  unsigned short x;        // 10-11
  unsigned short y;        // 12-13
  unsigned char  type;     // 14 - see above
  unsigned char  result;   // 15
} msg0x011F;

//--------------------------------------------------------------------
// 0x0120: Spell area disappeared
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned long  effect_id; // 02-05
} msg0x0120;

//--------------------------------------------------------------------
// 0x0121: Cart limit
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned short items;     // 02-03
  unsigned short itemsMax;  // 04-05
  unsigned long  weight;    // 06-09
  unsigned long  weightMax; // 10-13
} msg0x0121;

//--------------------------------------------------------------------
// 0x0122: List Equipments in cart
typedef struct
{
  unsigned short index;       // 00-01
  unsigned short type;        // 02-03
  unsigned char  category;    // 04
  unsigned char  identified;  // 05
  unsigned short equiptype;   // 06-07
  unsigned short equip_point; // 08-09
  unsigned char  element;     // 10
  unsigned char  refine;      // 11
//  unsigned short slot[4];     // 12-19
  union
  {
    unsigned short slot[4];

    struct
    {
      unsigned short flag;     // always = 0x00FF
      unsigned char  element;  // 1 - Water, 2 - Earth, 3 - Fire, 4 - Wind
      unsigned char  strength; // 5 - very strong, 10 - very very strong
      unsigned long  bsid;     // id of BS who build this weapon
    } smitten;

  } attributes;               // 12-19
} msg0x0122;

//--------------------------------------------------------------------
// 0x0123: List items in cart
typedef struct
{
  unsigned short index;   // 00-01
  unsigned short type;    // 02-03
  unsigned short unknown; // 04-05
  unsigned long  amount;  // 06-09
} msg0x0123;

//--------------------------------------------------------------------
// 0x0124: Item added to cart
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short index;    // 02-03
  unsigned long  amount;   // 04-07
  unsigned short type;     // 08-09
  unsigned short unknown;  // 10-11
  unsigned char  refine;   // 12
//  unsigned short slot[4];  // 13-20
  union
  {
    unsigned short slot[4];

    struct
    {
      unsigned short flag;     // always = 0x00FF
      unsigned char  element;  // 1 - Water, 2 - Earth, 3 - Fire, 4 - Wind
      unsigned char  strength; // 5 - very strong, 10 - very very strong
      unsigned long  bsid;     // id of BS who build this weapon
    } smitten;

  } attributes;               // 13-20
} msg0x0124;

//--------------------------------------------------------------------
// 0x0125: Item removed from cart
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short index;  // 02-03
  unsigned long  amount; // 04-07
} msg0x0125;

//--------------------------------------------------------------------
// 0x0126: Add item to cart
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short index;  // 02-03
  unsigned long  amount; // 04-07
} msg0x0126;

//--------------------------------------------------------------------
// 0x0127: Get item from cart
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short index;  // 02-03
  unsigned long  amount; // 04-07
} msg0x0127;

//--------------------------------------------------------------------
// 0x012C: Cannot add to cart
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned char  result;// 02 - 0 = cart overweight
} msg0x012C;

//--------------------------------------------------------------------
// 0x012D: Vending
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned short vacants; // 02-03 (number of items that can be sell)
} msg0x012D;

//--------------------------------------------------------------------
// 0x0131: Found vending store
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned long  store_id;    // 02-05
  unsigned char  title[36];   // 06-41
  unsigned char  unknown[44]; // 42-85
} msg0x0131;

//--------------------------------------------------------------------
// 0x0132: Item removed from vending store
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned long  itemID; // 02-05
} msg0x0132;

//--------------------------------------------------------------------
// 0x0133: List of item in vending store
typedef struct
{
  unsigned long  price;   // 00-03
  unsigned short amount;  // 04-05
  unsigned short index;   // 06-07
  unsigned char  type;    // 08
  unsigned long  itemID;  // 09-10
  unsigned char  identified; // 11
  unsigned char  unknown;    // 12
  unsigned char  custom;  // 13
//  unsigned short slot[4];    // 14-21
  union
  {
    unsigned short slot[4];

    struct
    {
      unsigned short flag;     // always = 0x00FF
      unsigned char  element;  // 1 - Water, 2 - Earth, 3 - Fire, 4 - Wind
      unsigned char  strength; // 5 - very strong, 10 - very very strong
      unsigned long  bsid;     // id of BS who build this weapon
    } smitten;

  } attributes;               // 14-21
} msg0x0133;

//--------------------------------------------------------------------
// 0x0136: List of item added to vending store - need to check, how different with 0x0133?
typedef msg0x0133 msg0x0136;

//--------------------------------------------------------------------
// 0x0137: Item sold
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short index;  // 02-03
  unsigned short amount; // 04-05
} msg0x0137;

//--------------------------------------------------------------------
// 0x0139: target locking failed
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned long  ID;    // 02-05
  unsigned short x;     // 06-07
  unsigned short y;     // 08-09
  unsigned short srcX;  // 10-11
  unsigned short srcY;  // 12-13
  unsigned short range; // 14-15
} msg0x0139;

//--------------------------------------------------------------------
// 0x013A: attack range
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned short range; // 02-03
} msg0x013A;

//--------------------------------------------------------------------
// 0x013B: Unequip arrow
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned short unknown; // 02-03
} msg0x013B;

//--------------------------------------------------------------------
// 0x013C: Equip Arrow
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned short index; // 02-03
} msg0x013C;

//--------------------------------------------------------------------
// 0x013D: HP/SP recovery
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short type;   // 02-03
  unsigned short amount; // 04-05
} msg0x013D;

//--------------------------------------------------------------------
// 0x013E: Spell casting
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned long  src_id;   // 02-05
  unsigned long  dst_id;   // 06-09 - 0 if casted on ground
  unsigned short x;        // 10-11 - 0 if casted on target
  unsigned short y;        // 12-13 - 0 if casted on target
  unsigned short skill_id; // 14-15
  unsigned long  element;  // 16-19
  unsigned long  delay;    // 20-23 - delay in ms
} msg0x013E;

//--------------------------------------------------------------------
// 0x0141: Bonus Status Information
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned short type;  // 02-03
  unsigned short unknown1; // 04-05
  unsigned short base;  // 06-07
  unsigned short unknown2; // 08-09
  unsigned short bonus;    // 10-11
  unsigned short unknown3; // 12-13
} msg0x0141;

//--------------------------------------------------------------------
// 0x0145: Display NPC picture
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned char  npcname[16]; // 02-17
  unsigned char  flag;     // 18
} msg0x0145;

//--------------------------------------------------------------------
// 0x0146: End talking to NPC
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned long  npcID;  // 02-05
} msg0x0146;

//--------------------------------------------------------------------
// 0x0147: Skill used
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned short skillID;  // 02-03
  unsigned char  unknown[35]; // 04-38
} msg0x0147;

//--------------------------------------------------------------------
// 0x0148: Unknown
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned long  id;      // 02-05
  unsigned short unknown; // 06-07
} msg0x0148;

//--------------------------------------------------------------------
// 0x0149: Align player
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned long  ID;     // 02-05
  unsigned char  alignment; // 06
} msg0x0149;

//--------------------------------------------------------------------
// 0x014B: Set manner point
typedef struct
{
  // the data size is proably just 4 bytes!!
  // check the packet log for the confirmation
  unsigned short cmd;      // 00-01
  unsigned char  value;    // 02
  unsigned char  name[24]; // 03-26
} msg0x014B;

//--------------------------------------------------------------------
// 0x014E: status of guild member
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned long  flag; // 02-05
} msg0x014E;

//--------------------------------------------------------------------
// 0x0152: Unknown
typedef struct
{
  // this message is variable length. Require to check the packet log
} msg0x0152;

//--------------------------------------------------------------------
// 0x0162: Guild information?
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short msglen; // 02-03
  unsigned char  data[]; // 04-
} msg0x0162;

//--------------------------------------------------------------------
// 0x016C: Guild Name
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned long  id;          // 02-05
  unsigned char  unknown[13]; // 06-18
  unsigned char  name[24];    // 19-42
} msg0x016C;

//--------------------------------------------------------------------
// 0x016D: Guild member online status
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned long  id;      // 02-05
  unsigned long  char_id; // 06-09
  unsigned long  status;  // 10-13
//  unsigned char  unknown[12]; // 02-13
} msg0x016D;

//--------------------------------------------------------------------
// 0x016F: Guild notice message
typedef struct
{
  unsigned short cmd;          // 00-01
  unsigned char  position[60]; // 02-61
  unsigned char  message[120]; // 62-181
} msg0x016F;

//--------------------------------------------------------------------
// 0x0177: List of items that possible to be identified
typedef struct
{
  unsigned short index; // 00-01
} msg0x0177;

//--------------------------------------------------------------------
// 0x0178: Use Item Appriasal
typedef struct
{
  unsigned short cmd;   // 00-01
  unsigned short index; // 02-03
} msg0x0178;

//--------------------------------------------------------------------
// 0x0179: Item identified
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned short index;    // 02-03
  unsigned char  unknown[8]; // 04-11
} msg0x0179;

//--------------------------------------------------------------------
// 0x017E: Guild Chat
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short msglen; // 02-03
  unsigned char  msg[];  // 04-
} msg0x017E;

//--------------------------------------------------------------------
// 0x017F: Chat message form guild
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short msglen; // 02-03
//  unsigned long  ID;  // 04-07
  unsigned char  msg[];  // 08-
} msg0x017F;

//--------------------------------------------------------------------
// 0x0180: Unknown
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned char  unknown[9]; // 02-10
} msg0x0180;

//--------------------------------------------------------------------
// 0x0182: Character attribute
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned long  acc_id;      // 02-05
  unsigned long  char_id;     // 06-09
  unsigned short hair_type;   // 10-11
  unsigned short hair_color;  // 12-13
  unsigned short sex;         // 14-15
  unsigned short job;         // 16-17
  unsigned short level;       // 18-19
  unsigned long  exp;         // 20-23
  unsigned long  online;      // 24-27
  unsigned long  position;    // 28-31
  unsigned char  unknown[50]; // 32-81
  unsigned char  name[24];    // 82-105
} msg0x0182;

//--------------------------------------------------------------------
// 0x0183: Unknown
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned char  unknown[13]; // 02-14
} msg0x0183;

//--------------------------------------------------------------------
// 0x0187: Got account ID
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned long  accountID; // 02-05
} msg0x0187;

//--------------------------------------------------------------------
// 0x018A: Request Terminating
//   type: Both
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned short unknown; // 02-03
} msg0x018A;

//--------------------------------------------------------------------
// 0x018B: Exit to windows (incoming)
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned short result; // 02-03
} msg0x018B;

//--------------------------------------------------------------------
// 0x0192: Unknown
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned long  id;          // 02-05
  unsigned short unknown;     // 06-07
  unsigned char  mapname[16]; // 08-23
} msg0x0192;

//--------------------------------------------------------------------
// 0x0194: Guild member connected
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned long  ID;       // 02-05
  unsigned char  name[24]; // 06-29
} msg0x0194;

//--------------------------------------------------------------------
// 0x0195: Character Information
typedef struct
{
  unsigned short cmd;           // 00-01
  unsigned long  ID;            // 02-05
  unsigned char  name[24];      // 06-29
  unsigned char  partyName[24]; // 30-53
  unsigned char  guildName[24]; // 54-77
  unsigned char  title[24];     // 78-101
} msg0x0195;

//--------------------------------------------------------------------
// 0x0196: Character special status
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned short code; // 02-03
  unsigned long  id;   // 04-07
  unsigned char  flag; // 08
} msg0x0196;

//--------------------------------------------------------------------
// 0x0199: start PVP Mode ?
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned short flag; // 02-03
} msg0x0199;

//--------------------------------------------------------------------
// 0x019A: PVP Rank/Num
typedef struct
{
  unsigned short cmd;      // 00-01
  unsigned long  id;       // 02-05
  unsigned long  rank;     // 06-09
  unsigned long  num;      // 10-13
} msg0x019A;

//--------------------------------------------------------------------
// 0x019B: Player gain level
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned long  ID;   // 02-05
  unsigned long  type; // 06-09
} msg0x019B;

//--------------------------------------------------------------------
// 0x01A2: Pet's name
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned char  name;    // 02-25
  unsigned char  unknown[9]; // 26-34
} msg0x01A2;

//--------------------------------------------------------------------
// 0x01A4: Pet spawned
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned char  type;    // 02
  unsigned long  ID;      // 03-06
  unsigned char  unknown[4]; // 07-10
} msg0x01A4;

//--------------------------------------------------------------------
// 0x01AA: Unknown
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned char  unknown[8];  // 02-09
} msg0x01AA;

//--------------------------------------------------------------------
// 0x01AB: Unknown
// 0x0000 (0000):ab 01 xx xx xx xx 04 00  c5 ff ff ff                 ..@>.... ....
typedef struct
{
  unsigned short cmd;        // 00-01
  unsigned long  id;         // 02-05
  unsigned char  unknown[6]; // 06-11
} msg0x01AB;
//--------------------------------------------------------------------
// 0x01AD: List of Materials for Arrow Crafting
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned short length;  // 02-03
  unsigned short items[];
} msg0x01AD;
//--------------------------------------------------------------------
// 0x01AE: Select Materials for Arrow Crafting
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned short materials; // 02-03
} msg0x01AE;
//--------------------------------------------------------------------
// 0x01B3: Display Animated NPC picture
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned char  npcname[16]; // 02-17
  unsigned char  unknown[48]; // 18-65
  unsigned char  flag;        // 66
} msg0x01B3;

//--------------------------------------------------------------------
// 0x01B5: Time Remaining
typedef struct
{
  unsigned short cmd;           // 00-01
  unsigned short remain_period; // 02-03
  unsigned char  unknown1[2];   // 04-05
  unsigned short remain_time;   // 06-07
  unsigned char  unknown2[10];  // 08-17
} msg0x01B5;
//--------------------------------------------------------------------
// 0x01B9: Player Reborn
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned long  id;      // 02-05
} msg0x01B9;

//--------------------------------------------------------------------
// 0x01C0: MVP
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned long  mvp_id; // 02-05
} msg0x01C0;

//--------------------------------------------------------------------
// 0x01C4: Item added to storage
typedef struct
{
  unsigned short cmd;        // 00-01
  unsigned short index;      // 02-03
  unsigned long  total;      // 04-07
  unsigned short type;       // 08-09

  unsigned char  unknown;    // 10

  unsigned char  identified; // 11
  unsigned char  element;    // 12
  unsigned char  refine;     // 13
//  unsigned short slot[4];    // 14-21
  union
  {
    unsigned short slot[4];

    struct
    {
      unsigned short flag;     // always = 0x00FF
      unsigned char  element;  // 1 - Water, 2 - Earth, 3 - Fire, 4 - Wind
      unsigned char  strength; // 5 - very strong, 10 - very very strong
      unsigned long  bsid;     // id of BS who build this weapon
    } smitten;

  } attributes;               // 14-21
} msg0x01C4;

//--------------------------------------------------------------------
// 0x01C5: Item added to cart?
//
// 00 01 02 03 04 05 06 07 08 09 10 11 12 13 14 15 16 17 18 19 20 21
// -----------------------------------------------------------------
// c5 01 5d 00 01 00 00 00 5b 02 02 01 00 00 00 00 00 00 00 00 00 00
//
// Add Old Blue Box - 603 (0x025B) 1 ea. into cart as item #93
typedef struct
{
  unsigned short cmd;          // 00-01
  unsigned short index;        // 02-03
  unsigned short amount;       // 04-05
  unsigned short unknown1;     // 06-07
  unsigned short itemtype;     // 08-09
  unsigned short unknown2;     // 10-11
  unsigned char  unknown3[10]; // 12-21
} msg0x01C5;

//--------------------------------------------------------------------
// 0x01C8: Use Item?
typedef struct
{
  unsigned short cmd;        // 00-01
  unsigned short index;      // 02-03
  unsigned short item_id;    // 04-05
  unsigned long  acc_id;     // 06-09
  unsigned short remaining;  // 10-11
  unsigned char  amount;     // 12
} msg0x01C8;

//--------------------------------------------------------------------
// 0x01C9: Skill Effect
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned long  effect_id;   // 02-05
  unsigned long  owner_id;    // 06-09
  unsigned short x;           // 10-11
  unsigned short y;           // 12-13
  unsigned char  unknown[83]; // 14-96
} msg0x01C9;

//--------------------------------------------------------------------
// 0x01CF: ???
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned long  id;          // 02-05
  unsigned char  unknown[22]; // 06-27
} msg0x01CF;

//--------------------------------------------------------------------
// 0x01D0: reborn type - same as 0x00b2?
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned long  id;   // 02-05
  unsigned short type; // 06-07
} msg0x01D0;

//--------------------------------------------------------------------
// 0x01D2: ???
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned long  id;      // 02-05
  unsigned long  unknown; // 06-09
} msg0x01D2;

//--------------------------------------------------------------------
// 0x01D6: map type
//
// maptype:
//  0 - town
//  1 - field
//  3 - dungeon
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned short maptype; // 02-03
} msg0x01D6;

//--------------------------------------------------------------------
// 0x01D7: equipment appearance
typedef struct
{
  unsigned short cmd;    // 00-01
  unsigned long  ID;     // 02-05
  unsigned char  equipping; // 06
  unsigned short item_id;   // 07-08
  unsigned short unknown;   // 09-10
} msg0x01D7;

//--------------------------------------------------------------------
// 0x01DB: Request for password encryption key
typedef struct
{
  unsigned short cmd;     // 00-01 << 0x01DB
} msg0x01DB;

//--------------------------------------------------------------------
// 0x01DC: Password encryption key
typedef struct
{
  unsigned short cmd;  // 00-01
  unsigned short msglen;  // 02-03
  unsigned char  key[16]; // 04-19
} msg0x01DC;

//--------------------------------------------------------------------
// 0x01DD: Encrypted login packet
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned long  patch;   // 02-05
  unsigned char  login[24];  // 06-29
  unsigned char  passwd[17]; // 30-46
} msg0x01DD;

//--------------------------------------------------------------------
// 0x01DE: Attack with skill
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned short skill_id;  // 02-03
  unsigned long  src_id;    // 04-07
  unsigned long  dst_id;    // 08-11

  unsigned long  tick;      // 12-15
  unsigned long  src_speed; // 16-19
  unsigned long  dst_speed; // 20-23

  unsigned long  damage;    // 24-27
  unsigned short level;     // 28-29
  unsigned short hitnum;    // 30-31
  unsigned char  type;      // 32    - see 0x0114
} msg0x01DE;

//--------------------------------------------------------------------
// 0x01E1: ??
typedef struct
{
  unsigned short cmd;     // 00-01
  unsigned long  id;      // 02-05
  unsigned short unknown; // 06-07
} msg0x01E1;

//--------------------------------------------------------------------
// 0x01E6: ??
typedef struct
{
  unsigned short cmd;         // 00-01
  unsigned char  unknown[24]; // 02-25
} msg0x01E6;

//--------------------------------------------------------------------
// 0x01EE: Item in the Inventory
typedef struct
{
  unsigned short index;       // 00-01
  unsigned short type;        // 02-03
  unsigned char  category;    // 04
  unsigned char  identified;  // 05
  unsigned short amount;      // 06-07
  unsigned char  unknown[10]; // 08-17
} msg0x01EE;

//--------------------------------------------------------------------
// 0x01EF: Item in cart?
typedef struct
{
  unsigned short index;       // 00-01
  unsigned short type;        // 02-03
  unsigned char  category;    // 04
  unsigned char  identified;  // 05
  unsigned short amount;      // 06-07
  unsigned char  unknown[10]; // 08-17
} msg0x01EF;

//--------------------------------------------------------------------
// 0x01F0: Item in storage (extended)
typedef struct
{
  unsigned short index;       // 00-01
  unsigned short type;        // 02-03
  unsigned char  category;    // 04
  unsigned char  identified;  // 05
  unsigned long  amount;      // 06-09
  unsigned char  unknown2[8]; // 10-17
} msg0x01F0;

//--------------------------------------------------------------------
// 0x01F2: Something about guild?
typedef struct
{
  unsigned short cmd;       // 00-01
  unsigned long  guild_id;  // 02-05
  unsigned long  member_id; // 06-09
  unsigned short unknown1;  // 10-11
  unsigned short unknown2;  // 12-13
  unsigned short unknown3;  // 14-15
  unsigned short unknown4;  // 16-17
  unsigned short unknown5;  // 18-19
} msg0x01F2;

//--------------------------------------------------------------------
// 0x01F4: Unknown
typedef struct
{
  unsigned short cmd;        // 00-01
  unsigned char  unknown[5]; // 02-06
} msg0x01F4;

//--------------------------------------------------------------------
// 0x0200: Pre-login?
typedef struct
{
  unsigned short cmd;          // 00-01
  unsigned char  username[24]; // 02-25

} msg0x0200;

#pragma pack()

#ifdef __cplusplus
}
#endif

#endif // _PACKET_H

