Attribute VB_Name = "md_crc32"
'// At top level of a module, always include to be sure that all
'variables have the right type
Option Explicit
Option Compare Text

'// Then declare this array variable Crc32Table
Private Crc32Table(255) As Long

'// Then all we have to do is write public functions like
'these...
Public Function InitCrc32(Optional ByVal Seed As Long = _
   &HEDB88320, Optional ByVal Precondition As _
   Long = &HFFFFFFFF) As Long

   '// Declare counter variable iBytes,
   'counter variable iBits,
    'value variables lCrc32 and lTempCrc32
   
Dim iBytes As Integer, iBits As Integer, lCrc32 As Long
Dim lTempCrc32 As Long

   '// Turn on error trapping
   On Error Resume Next

   '// Iterate 256 times
   For iBytes = 0 To 255

      '// Initiate lCrc32 to counter variable
      lCrc32 = iBytes

      '// Now iterate through each bit in counter byte
      For iBits = 0 To 7
         '// Right shift unsigned long 1 bit
         lTempCrc32 = lCrc32 And &HFFFFFFFE
         lTempCrc32 = lTempCrc32 \ &H2
         lTempCrc32 = lTempCrc32 And &H7FFFFFFF

         '// Now check if temporary is less than zero and then
         'mix Crc32 checksum with Seed value
         If (lCrc32 And &H1) <> 0 Then
            lCrc32 = lTempCrc32 Xor Seed
         Else
            lCrc32 = lTempCrc32
         End If
      Next

      '// Put Crc32 checksum value in the holding array
      Crc32Table(iBytes) = lCrc32
   Next

   '// After this is done, set function value to the
   'precondition value
   InitCrc32 = Precondition

End Function

'// The function above is the initializing function, now
'we have to write the computation function
Public Function AddCrc32(ByVal Item As String, _
  ByVal Crc32 As Long) As Long

   '// Declare following variables
   Dim bCharValue As Byte, iCounter As Integer, lIndex As Long
   Dim lAccValue As Long, lTableValue As Long

   '// Turn on error trapping
   On Error Resume Next

   '// Iterate through the string that is to be checksum-computed
   For iCounter = 1 To Len(Item)

      '// Get ASCII value for the current character
      bCharValue = Asc(Mid$(Item, iCounter, 1))

      '// Right shift an Unsigned Long 8 bits
      lAccValue = Crc32 And &HFFFFFF00
      lAccValue = lAccValue \ &H100
      lAccValue = lAccValue And &HFFFFFF

      '// Now select the right adding value from the
       'holding table
      lIndex = Crc32 And &HFF
      lIndex = lIndex Xor bCharValue
      lTableValue = Crc32Table(lIndex)

      '// Then mix new Crc32 value with previous
      'accumulated Crc32 value
      Crc32 = lAccValue Xor lTableValue
   Next

   '// Set function value the the new Crc32 checksum
   AddCrc32 = Crc32

End Function

'// At last, we have to write a function so that we
'can get the Crc32 checksum value at any time
Public Function GetCrc32(ByVal Crc32 As Long) As Long
   '// Turn on error trapping
   On Error Resume Next

   '// Set function to the current Crc32 value
   GetCrc32 = Crc32 Xor &HFFFFFFFF

End Function

Function GetCRC(ByVal Crc32 As String) As Long
    Dim lCrc32Value As Long
    lCrc32Value = InitCrc32()
    lCrc32Value = AddCrc32(Crc32, lCrc32Value)
    GetCRC = GetCrc32(lCrc32Value)
End Function
