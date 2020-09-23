<div align="center">

## DecryptBase64String


</div>

### Description

This one is to show how to DECODE Base64. Base64 is used to encode Mime Attachements. This not a complet Mime Decoder, this routine should just show how to build one!

By the way the hole programm, which is able to decode Mime will follow...
 
### More Info
 
Base64 String

I build an example programm, just look!

Copy the text below and save it as DecodeMime.frm and compile it.

Binary

I just add a little errorcheck but I'm not sure if this will be enough!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sebastian Fahrenkrog](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sebastian-fahrenkrog.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sebastian-fahrenkrog-decryptbase64string__1-2387/archive/master.zip)





### Source Code

```
'Copy the part below and paste it into the Notepad
'and save it as DecodeMime.frm
'-------------------------8< Cut here ----------------------------------------
VERSION 5.00
Begin VB.Form Form1
  BorderStyle   =  4 'Festes Werkzeugfenster
  Caption     =  "Base64 Decode Example"
  ClientHeight  =  2205
  ClientLeft   =  45
  ClientTop    =  300
  ClientWidth   =  6000
  LinkTopic    =  "Form1"
  MaxButton    =  0  'False
  MinButton    =  0  'False
  ScaleHeight   =  2205
  ScaleWidth   =  6000
  ShowInTaskbar  =  0  'False
  StartUpPosition =  2 'Bildschirmmitte
  Begin VB.CommandButton Decode
   Caption     =  "Decode"
   Height     =  495
   Left      =  1800
   TabIndex    =  2
   Top       =  1560
   Width      =  1815
  End
  Begin VB.TextBox Binary
   Height     =  285
   Left      =  240
   TabIndex    =  1
   Top       =  1080
   Width      =  5295
  End
  Begin VB.TextBox Base64
   Height     =  285
   Left      =  240
   TabIndex    =  0
   Text      =  "N6iOK/rfOyMWYyJ5EVHoLdFLty707JuWNhr5aCI8YGsOIDQTLdv7sQ=="
   Top       =  480
   Width      =  5295
  End
  Begin VB.Label Label2
   Caption     =  "Binarys:"
   Height     =  255
   Left      =  240
   TabIndex    =  4
   Top       =  840
   Width      =  735
  End
  Begin VB.Label Label1
   Caption     =  "Base64:"
   Height     =  255
   Left      =  240
   TabIndex    =  3
   Top       =  240
   Width      =  735
  End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************
'This is the Base64 Decode Example and show you how to
'decode Base64!
'
'At the moment I'm to laszy to write a hole programm to
'decrypt Mime Attachements, so if you want you can take
'this example of how to do it right and write you own
'routine! You have to write a few routines to find the
'specific Mime headers. If you want to know more about
'this, send me an E-Mail...
'
'E-mail: galgen@wtal.de
'*********************************************************
Private Function Base64Decode(Basein As String) As String
Dim counter As Integer
Dim Temp As String
'For the dec. Tab
Dim DecodeTable As Variant
Dim Out(2) As Byte
Dim inp(3) As Byte
'DecodeTable holds the decode tab
DecodeTable = Array("255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "62", "255", "255", "255", "63", "52", "53", "54", "55", "56", "57", "58", "59", "60", "61", "255", "255", "255", "64", "255", "255", "255", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", _
"18", "19", "20", "21", "22", "23", "24", "25", "255", "255", "255", "255", "255", "255", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255" _
, "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255")
'Reads 4 Bytes in and decrypt them
For counter = 1 To Len(Basein) Step 4
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!IF YOU WANT YOU CAN ADD AN ERRORCHECK:         !
'!If DecodeTable()=255 Then Error!            !
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'4 Bytes in -> 3 Bytes out
inp(0) = DecodeTable(Asc(Mid$(Basein, counter, 1)))
inp(1) = DecodeTable(Asc(Mid$(Basein, counter + 1, 1)))
inp(2) = DecodeTable(Asc(Mid$(Basein, counter + 2, 1)))
inp(3) = DecodeTable(Asc(Mid$(Basein, counter + 3, 1)))
Out(0) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
Out(1) = ((inp(1) And &HF) * 16) Or ((inp(2) \ 4) And &HF)
Out(2) = ((inp(2) And &H3) * 64) Or inp(3)
'* look for "=" symbols
If inp(2) = 64 Then
  'If there are 2 characters left -> 1 binary out
  Out(0) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
  Temp = Temp & Chr(Out(0) And &HFF)
ElseIf inp(3) = 64 Then
  'If there are 3 characters left -> 2 binaries out
  Out(0) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
  Out(1) = ((inp(1) And &HF) * 16) Or ((inp(2) \ 4) And &HF)
  Temp = Temp & Chr(Out(0) And &HFF) & Chr(Out(1) And &HFF)
Else 'Return three Bytes
  Temp = Temp & Chr(Out(0) And &HFF) & Chr(Out(1) And &HFF) & Chr(Out(2) And &HFF)
End If
Next
Base64Decode = Temp
End Function
'**********************************************************
Private Sub Decode_Click()
'Base64 needs x * 4 Bytes to work...
If Base64 <> "" And (Len(Base64) Mod 4) = 0 Then
Binary.Text = Base64Decode(Base64.Text)
End If
End Sub
```

