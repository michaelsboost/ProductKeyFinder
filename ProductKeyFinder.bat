::' Windows RT 8.0 Product Key Dumper by Myria of xda-developers.com
::' Original Windows 8.0 VBScript by janek2012 of mydigitallife.info
::' Batch+VBScript hybrid trick by dbenham of stackoverflow.com
::' Fix for keys starting with N by Osprey00 of xda-developers.com
::'
::' Windows RT doesn't let unsigned VBScript use WScript.Shell, which is
::' required in order to read the registry in VBScript.  So instead, we
::' have a batch file call reg.exe to do the registry lookup for us, then
::' execute the VBScript code.  Might as well do things this way, since
::' it would really suck to write this math in batch...

::' --- Batch portion ---------
rem^ &@echo off
rem^ &call :'sub
::' If we were run from double-clicking in Explorer, pause.
rem^ &if %0 == "%~0" pause
rem^ &exit /b 0

:'sub
::' Read the registry key into VBScript's stdin.
rem^ &("%SystemRoot%\System32\reg.exe" query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v DigitalProductId | cscript //nologo //e:vbscript "%~f0")
::'rem^ &echo end batch
rem^ &exit /b 0

'----- VBS portion ------------
'WScript.Echo "begin VBS"

' Get registry data that was piped in
RegData = ""
Do While Not WScript.StdIn.AtEndOfStream
    RegData = RegData & WScript.StdIn.ReadAll
Loop

' Remove any carriage returns
RegData = Replace(RegData, ChrW(13), "")

' Split into lines
RegLines = Split(RegData, ChrW(10))

' Sanity checking on data
If (RegLines(0) <> "") Or (RegLines(1) <> "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion") Then
    WScript.Echo "Got invalid header trying to run reg.exe"
    WScript.Quit(1)
End If

If Left(RegLines(2), 38) <> "    DigitalProductId    REG_BINARY    " Then
    WScript.Echo "Got invalid value list trying to run reg.exe"
    WScript.Quit(1)
End If

' Get hex string
HexString = Mid(RegLines(2), 39)
If (Len(HexString) Mod 2) <> 0 Then
    WScript.Echo "Got an odd number of hex digits in REG_BINARY data"
    WScript.Quit(1)
End If

' Convert to byte array
Dim ByteArray()
ReDim ByteArray((Len(HexString) / 2) - 1)  ' VBScript is just weird with array dimensions >.<

For i = 0 To (Len(HexString) - 2) Step 2
    ByteArray(i / 2) = CInt("&H" + Mid(HexString, i + 1, 2))
Next

Key = ConvertToKey(ByteArray)
WScript.Echo Key

' janek2012's magic decoding function
Function ConvertToKey(Key)
    Const KeyOffset = 52 ' Offset of the first byte of key in DigitalProductId - helps in loops
    isWin8 = (Key(66) \ 8) And 1 ' Check if it's Windows 8 here...
    Key(66) = (Key(66) And &HF7) Or ((isWin8 And 2) * 4) ' Replace 66 byte with logical result
    Chars = "BCDFGHJKMPQRTVWXY2346789" ' Characters used in Windows key
    ' Standard Base24 decoding...
    For i = 24 To 0 Step -1
        Cur = 0
        For X = 14 To 0 Step -1
            Cur = Cur * 256
            Cur = Key(X + KeyOffset) + Cur
            Key(X + KeyOffset) = (Cur \ 24)
            Cur = Cur Mod 24
        Next
        KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
        Last = Cur
    Next
    ' If it's Windows 8, put "N" in the right place
    If (isWin8 = 1) Then
        keypart1 = Mid(KeyOutput, 2, Cur)
        insert = "N"
        KeyOutput = keypart1 & insert & Mid(KeyOutput, Cur + 2)
    End If
    ' Divide keys to 5-character parts
    a = Mid(KeyOutput, 1, 5)
    b = Mid(KeyOutput, 6, 5)
    c = Mid(KeyOutput, 11, 5)
    d = Mid(KeyOutput, 16, 5)
    e = Mid(KeyOutput, 21, 5)
    ' And join them again adding dashes
    ConvertToKey = a & "-" & b & "-" & c & "-" & d & "-" & e
    ' The result of this function is now the actual product key
End Function