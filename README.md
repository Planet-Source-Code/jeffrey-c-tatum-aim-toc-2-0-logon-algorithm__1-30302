<div align="center">

## AIM Toc 2\.0 Logon Algorithm


</div>

### Description

This will create the 8 or 9 digit number that goes along with the log on packet for the newest version of the AIM TOC Protocol (version 2.0). The number that it makes, is based on the screen name and password. I guess it was a security feature added in, to prevent users from making AIM Crackers.
 
### More Info
 
Since I'm sure many people don't actualy know how to connect to the AIM server, I will give an example with this code. This code will successfully sign you on to the AIM server using TOC 2.0 protocol.

First, create a form. Add the following controls:

Text1 = The Screen Name

Text2 = The Password

Text3 = The Incomming Text

Winsock1 = Microsoft Winsock Control

Command1 = The Connect Button

Once the form is created, copy and paste all of the following code into the declarations part of your form. Once you copy it all into your form, simply run the application, type in your aim screen name into Text1, your password into Text2, and click the Command1 button to connect.

Returns a number based on the screen name and password, to enable the user to log on to the AIM Toc 2.0 server.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeffrey C\. Tatum](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeffrey-c-tatum.md)
**Level**          |Beginner
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeffrey-c-tatum-aim-toc-2-0-logon-algorithm__1-30302/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
Winsock1.Connect "aimexpress.oscar.aol.com", 5190
End Sub
Function AIM_Algorithum(ByVal sUser As String, ByVal sPass As String) As String
'This is the code that generates the 8 or 9 digit number on the end of
'the logon packet. Uses the Screen Name and Password to make it
Dim sUserChar As Long, sVar As Long
  DoEvents: sUser = Left(LCase(sUser), 1)
  DoEvents: sUserChar = Int(Asc(sUser) - 96)
  DoEvents: sVar = Int(sUserChar * 7696) + 738816
  DoEvents: sBase = Int(sUserChar * 746512)
  DoEvents: sVal = Int(Asc(Left(LCase(sPass), 1)) - 96) * sVar
  AIM_Algorithum = Int(Int(sVal) - sVar) + Int(sBase + 71665152)
End Function
Function AIM_EncryptPW(ByVal sPass As String) As String
'This will take the password, and encrypt it using the word "Tic/Toc"
Dim vTable() As Variant, sString As String
Dim sLoop As Long, sHex As String
vTable = Array("84", "105", "99", "47", "84", "111", "99")
sString = "0x"
For sLoop = 0 To Len(sPass) - 1
  sHex = Hex(Asc(Mid(sPass, sLoop + 1, 1)) Xor CLng(vTable(sLoop Mod 7)))
  If CLng("&H" & sHex) < 16 Then
    sString = sString & "0"
  End If
  sString = sString & sHex
Next
AIM_EncryptPW = LCase(sString)
End Function
Private Sub Form_Load()
End Sub
Private Sub Winsock1_Connect()
Winsock1.SendData "FLAPON" & vbCrLf & vbCrLf
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Text As String
'Get the data from the server.
Winsock1.GetData Text, vbString
'Place all the incomming text, into text3, so you
'can see what all is going on. I replaced character
'0, with Ø so that you can see the text with the
'null character.
Text3 = Text3 & vbCrLf & Replace(Text, Chr(0), "Ø")
'If the second character is character 1, it means it
'wants the log on information. You will only get that
'character at log on, and never again once you're
'connected.
If Asc(Mid(Text, 2, 1)) = 1 Then
  'Send the log on information
  Winsock1.SendData Chr(42) & Chr(1) & Chr(1) & Chr(0) & Chr(0) & Chr(8 + Len(Text1)) & Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(0) & Chr(1) & Chr(0) & Chr(Len(Text1)) & Trim(Text1)
  Winsock1.SendData Chr(42) & Chr(2) & Chr(1) & Chr(1) & Chr(0) & Chr(Len(Text1) + Len(AIM_EncryptPW(Trim(Text2))) + 90) & "toc2_signon login.oscar.aol.com 29999 " & Trim(Text1) & " " & AIM_EncryptPW(Trim(Text2)) & " english-US " & Chr(34) & "TIC:\$Revision: 1.83 \$" & Chr(34) & " 160 " & AIM_Algorithum(Text1, Text2) & Chr(0)
End If
End Sub
```

