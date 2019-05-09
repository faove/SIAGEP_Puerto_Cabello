Attribute VB_Name = "IDAutomationVBA"
'Option Compare Database

'*********************************************************************
'*
'*  Visual Basic / VBA Functions for Bar Code Fonts 2.11
'*  Copyright, IDAutomation.com, Inc. 2000. All rights reserved.
'*
'*  Visit http://www.BizFonts.com/vba/ for more information.
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid Multi-user, Corporate or Distribution
'*  license from IDAutomation.com, Inc. for the associated font and
'*  the copyright notices are not removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a distribution license.
'*
'*  To find a particular function, search the text for it.
'*
'*  HOW TO USE IN VISUAL BASIC:
'*  For best results, just insert this file as a VB module
'*  and access the function in your application as necessary.
'*
'*  To do this, rename this file to IDAutomationVBA.bas
'*  Then, in your VB project, choose Project - Add File and select this file.
'*  After that, you can access the function you need.
'*  Example: Printer.Print Code128b("123456789")
'*
'*********************************************************************

'Attribute VB_Name is only used for some Microsoft applications
'remark this out if you use this as LotusScript

'START OF DECLARACTIONS
Private i As Integer
Private f As Integer
Private DataToPrint As String
Private DataToEncode As String
Private DataToFormat As String
Private OnlyCorrectData As String
Private Printable_string As String
Private Encoding As String
Private weightedTotal As Long
Private WeightValue As Integer
Private CurrentValue As Long
Private CheckDigitValue As Integer
Private Factor As Integer
Private CheckDigit As Integer
Private CurrentEncoding As String
Private NewLine As String
Private Msg As String
Private CurrentChar As String
Private CurrentCharNum As Integer
Private C128_StartA As String
Private C128_StartB As String
Private C128_StartC As String
Private C128_Stop As String
Private C128Start As String
Private C128_CheckDigit As String
Private StartCode As String
Private StopCode As String
Private Fnc1 As String
Private LeadingDigit As Integer
Private EAN2AddOn As String
Private EAN5AddOn As String
Private EANAddOnToPrint As String
Private HumanReadableText As String
Private StringLength As Integer
'END OF DECLARACTIONS


Public Function Postnet(DataToEncode As String, ReturnType As Integer) As String
' Copyright © IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com or http://www.IDautomation.com
'
' The purpose of this code is to calculate the POSTNET barcode
' Enter all the numbers without dashes
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
     DataToEncode = OnlyCorrectData
'<<<< Calculate Check Digit >>>>
     weightedTotal = 0
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Get the value of each number
          CurrentCharNum = Mid(DataToEncode, i, 1)
    'add the values together
          weightedTotal = weightedTotal + CurrentCharNum
     Next i
'Find the CheckDigit by finding the number + weightedTotal that = a multiple of 10
'divide by 10, get the remainder and subtract from 10
     i = (weightedTotal Mod 10)
     If i <> 0 Then
          CheckDigit = (10 - i)
     Else
          CheckDigit = 0
     End If
'Get Printable String
     DataToPrint = DataToEncode
'ReturnType 0 returns data formatted to the barcode font
     If ReturnType = 0 Then Postnet = "(" & DataToPrint & CheckDigit & ")" & " "
'ReturnType 1 returns data formatted for human readable text
     If ReturnType = 1 Then Postnet = DataToPrint & CheckDigit
'ReturnType 2 returns the  check digit for the data supplied
     If ReturnType = 2 Then Postnet = STR$(CheckDigit)
End Function

Public Function Code128(DataToFormat As String, ReturnType As Integer) As String
'
' Copyright © IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.IDAutomation.com
'
' You may use our source code in your applications only if you are using barcode fonts created by IDautomation.com, Inc.
' and you do not remove the copyright notices in the source code.
'
' The purpose of this code is to calculate the Code 128 barcode for ANY character set
'
' Encode UCC/EAN 128 by inserting ASCII 202 into the string to encode
'
' You MUST use the fully functional Code 128 (dated 12/2000 or later)
' font for this code to create and print a proper barcode
'
    DataToPrint = ""
    DataToFormat = RTrim(LTrim(DataToFormat))
    C128_StartA = Chr(203)
    C128_StartB = Chr(204)
    C128_StartC = Chr(205)
    C128_Stop = Chr(206)
'Here we select character set A, B or C for the START character
    StringLength = Len(DataToFormat)
    CurrentCharNum = Asc(Mid(DataToFormat, 1, 1))
    If CurrentCharNum < 32 Then C128Start = C128_StartA
    If CurrentCharNum > 31 And CurrentCharNum < 127 Then C128Start = C128_StartB
    If ((StringLength > 4) And IsNumeric(Mid(DataToFormat, 1, 4))) Then C128Start = C128_StartC
'202 is the FNC1, with this Start C is mandatory
    If CurrentCharNum = 202 Then C128Start = C128_StartC
    If C128Start = Chr(203) Then CurrentEncoding = "A"
    If C128Start = Chr(204) Then CurrentEncoding = "B"
    If C128Start = Chr(205) Then CurrentEncoding = "C"
    For i = 1 To StringLength
    'check for FNC1 in any set
        If ((Mid(DataToFormat, i, 1)) = Chr(202)) Then
            DataToEncode = DataToEncode & Chr(202)
    'check for switching to character set C
        ElseIf ((i < StringLength - 2) And (IsNumeric(Mid(DataToFormat, i, 1))) And (IsNumeric(Mid(DataToFormat, i + 1, 1))) And (IsNumeric(Mid(DataToFormat, i, 4)))) Or ((i < StringLength) And (IsNumeric(Mid(DataToFormat, i, 1))) And (IsNumeric(Mid(DataToFormat, i + 1, 1))) And (CurrentEncoding = "C")) Then
        'switch to set C if not already in it
            If CurrentEncoding <> "C" Then DataToEncode = DataToEncode & Chr(199)
            CurrentEncoding = "C"
            CurrentChar = (Mid(DataToFormat, i, 2))
            CurrentValue = CInt(CurrentChar)
        'set the CurrentValue to the number of String CurrentChar
            If (CurrentValue < 95 And CurrentValue > 0) Then DataToEncode = DataToEncode & Chr(CurrentValue + 32)
            If CurrentValue > 94 Then DataToEncode = DataToEncode & Chr(CurrentValue + 100)
            If CurrentValue = 0 Then DataToEncode = DataToEncode & Chr(194)
            i = i + 1
    'check for switching to character set A
        ElseIf (i <= StringLength) And ((Asc(Mid(DataToFormat, i, 1)) < 31) Or ((CurrentEncoding = "A") And (Asc(Mid(DataToFormat, i, 1)) > 32 And (Asc(Mid(DataToFormat, i, 1))) < 96))) Then
        'switch to set A if not already in it
            If CurrentEncoding <> "A" Then DataToEncode = DataToEncode & Chr(201)
            CurrentEncoding = "A"
        'Get the ASCII value of the next character
            CurrentCharNum = (Asc(Mid(DataToFormat, i, 1)))
            If CurrentCharNum = 32 Then
                DataToEncode = DataToEncode & Chr(194)
            ElseIf CurrentCharNum < 32 Then
                DataToEncode = DataToEncode & Chr(CurrentCharNum + 96)
            ElseIf CurrentCharNum > 32 Then
                DataToEncode = DataToEncode & Chr(CurrentCharNum)
            End If
    'check for switching to character set B
        ElseIf (i <= StringLength) And ((Asc(Mid(DataToFormat, i, 1))) > 31 And (Asc(Mid(DataToFormat, i, 1)))) < 127 Then
        'switch to set B if not already in it
            If CurrentEncoding <> "B" Then DataToEncode = DataToEncode & Chr(200)
            CurrentEncoding = "B"
        'Get the ASCII value of the next character
            CurrentCharNum = (Asc(Mid(DataToFormat, i, 1)))
            If CurrentCharNum = 32 Then
                DataToEncode = DataToEncode & Chr(194)
            Else
                DataToEncode = DataToEncode & Chr(CurrentCharNum)
            End If
        End If
    Next i
    
    HumanReadableText = ""
'FORMAT TEXT FOR AIs
    StringLength = Len(DataToFormat)
    For i = 1 To StringLength
    'Get ASCII value of each character
        CurrentCharNum = Asc(Mid(DataToFormat, i, 1))
    'Check for FNC1
        If ((i < StringLength - 2) And (CurrentCharNum = 202)) Then
        'It appears that there is an AI
        'Get the value of each number pair (ex: 5 and 6 = 5*10+6 =56)
            CurrentChar = (Mid(DataToFormat, i + 1, 2))
            CurrentCharNum = CInt(CurrentChar)
        'Is 4 digit AI?
            If ((i < StringLength - 4) And ((CurrentCharNum <= 81 And CurrentCharNum >= 80) Or (CurrentCharNum <= 34 And CurrentCharNum >= 31))) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, i + 1, 4)) & ") "
                i = i + 4
        'Is 3 digit AI?
            ElseIf ((i < StringLength - 3) And ((CurrentCharNum <= 49 And CurrentCharNum >= 40) Or (CurrentCharNum <= 25 And CurrentCharNum >= 23))) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, i + 1, 3)) & ") "
                i = i + 3
        'Is 2 digit AI?
            ElseIf ((CurrentCharNum <= 30 And CurrentCharNum >= 0) Or (CurrentCharNum <= 99 And CurrentCharNum >= 90)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, i + 1, 2)) & ") "
                i = i + 2
            End If
        ElseIf (Asc(Mid(DataToFormat, i, 1)) < 32) Then
            HumanReadableText = HumanReadableText & " "
        ElseIf ((Asc(Mid(DataToFormat, i, 1)) > 31) And (Asc(Mid(DataToFormat, i, 1)) < 128)) Then
            HumanReadableText = HumanReadableText & Mid(DataToFormat, i, 1)
        End If
    Next i
    DataToFormat = ""
    
'<<<< Calculate Modulo 103 Check Digit >>>>
'Set WeightedTotal to the value of the start character
    weightedTotal = (Asc(C128Start) - 100)
    StringLength = Len(DataToEncode)
    For i = 1 To StringLength
    'Get the ASCII value of each character
        CurrentCharNum = (Asc(Mid(DataToEncode, i, 1)))
    'Get the Code 128 value of CurrentChar according to chart
        If CurrentCharNum < 135 Then CurrentValue = CurrentCharNum - 32
        If CurrentCharNum > 134 Then CurrentValue = CurrentCharNum - 100
        If CurrentCharNum = 194 Then CurrentValue = 0
    'multiply by the weighting character
        CurrentValue = CurrentValue * i
    'add the values together
        weightedTotal = weightedTotal + CurrentValue
    Next i
'divide the WeightedTotal by 103 and get the remainder, this is the CheckDigitValue
    CheckDigitValue = (weightedTotal Mod 103)
'Now that we have the CheckDigitValue, find the corresponding ASCII character from the table
    If CheckDigitValue < 95 And CheckDigitValue > 0 Then C128_CheckDigit = Chr(CheckDigitValue + 32)
    If CheckDigitValue > 94 Then C128_CheckDigit = Chr(CheckDigitValue + 100)
    If CheckDigitValue = 0 Then C128_CheckDigit = Chr(194)
'Check for spaces or "00" and print ASCII 194 instead
'place changes in DataToPrint
    StringLength = Len(DataToEncode)
    For i = 1 To StringLength
        CurrentChar = Mid(DataToEncode, i, 1)
        If CurrentChar = " " Then CurrentChar = Chr(194)
        DataToPrint = DataToPrint & CurrentChar
    Next i
'Get Printable String
    Printable_string = C128Start & DataToPrint & C128_CheckDigit & C128_Stop & " "
    DataToEncode = ""
    DataToPrint = ""
'ReturnType 0 returns data formatted to the barcode font
    If ReturnType = 0 Then Code128 = Printable_string
'ReturnType 1 returns data formatted for human readable text
    If ReturnType = 1 Then Code128 = HumanReadableText
'ReturnType 2 returns the check digit for the data supplied
    If ReturnType = 2 Then Code128 = C128_CheckDigit
    
End Function

Public Function Code128a(DataToEncode As String) As String
'
' This module is Copyright, IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com or http://www.IDautomation.com
'
' The purpose of this code is to print the Code 128 barcode from character set A
' Use the characters from set B to print characters not on the keyboard
' The scanner will scan characters from set A
'
' You MUST use the fully functional Code 128 (dated 12/2000 or later)
' font for this code to create and print a proper barcode
'
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
     C128_StartA = Chr(203)
     C128_StartB = Chr(204)
     C128_StartC = Chr(205)
     C128_Stop = Chr(206)
'Here we select character set A
     C128Start = C128_StartA
'<<<< Calculate Modulo 103 Check Digit >>>>
'Set WeightedTotal to the value of the start character
     weightedTotal = (Asc(C128Start) - 100)
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Get the ASCII value of each character
          CurrentCharNum = (Asc(Mid(DataToEncode, i, 1)))
    'Get the Code 128 value of CurrentChar according to chart
          If CurrentCharNum < 135 Then CurrentValue = CurrentCharNum - 32
          If CurrentCharNum > 134 Then CurrentValue = CurrentCharNum - 100
    'multiply by the weighting character
          CurrentValue = CurrentValue * i
    'add the values together
          weightedTotal = weightedTotal + CurrentValue
     Next i
'divide the WeightedTotal by 103 and get the remainder, this is the CheckDigitValue
     CheckDigitValue = (weightedTotal Mod 103)
'Now that we have the CheckDigitValue, find the corresponding ASCII character from the table
     If CheckDigitValue < 95 And CheckDigitValue > 0 Then C128_CheckDigit = Chr(CheckDigitValue + 32)
     If CheckDigitValue > 94 Then C128_CheckDigit = Chr(CheckDigitValue + 100)
     If CheckDigitValue = 0 Then C128_CheckDigit = Chr(194)
'Check for spaces or "00" and print ASCII 194 instead
'place changes in DataToPrint
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
          CurrentChar = Mid(DataToEncode, i, 1)
          If CurrentChar = " " Then CurrentChar = Chr(194)
          DataToPrint = DataToPrint & CurrentChar
     Next i
'Get PrintableString
     Printable_string = C128Start & DataToPrint & C128_CheckDigit & C128_Stop & " "
'Return the PrintableString
     Code128a = Printable_string
End Function



Public Function Code128b(DataToEncode As String) As String
'
' Copyright © IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com or http://www.IDautomation.com
'
' You may use our source code in your applications only if you are using barcode fonts created by IDautomation.com, Inc.
' and you do not remove the copyright notices in the source code.
'
' The purpose of this code is to calculate the Code 128 barcode from character set B
'
' You MUST use the fully functional Code 128 (dated 12/2000 or later)
' font for this code to create and print a proper barcode
'
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
     C128_StartA = Chr(203)
     C128_StartB = Chr(204)
     C128_StartC = Chr(205)
     C128_Stop = Chr(206)
'Here we select character set A or B but not C
     C128Start = C128_StartB
'<<<< Calculate Modulo 103 Check Digit >>>>
'Set WeightedTotal to the value of the start character
     weightedTotal = (Asc(C128Start) - 100)
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Get the ASCII value of each character
          CurrentCharNum = (Asc(Mid(DataToEncode, i, 1)))
    'Get the Code 128 value of CurrentChar according to chart
          If CurrentCharNum < 135 Then CurrentValue = CurrentCharNum - 32
          If CurrentCharNum > 134 Then CurrentValue = CurrentCharNum - 100
    'multiply by the weighting character
          CurrentValue = CurrentValue * i
    'add the values together
          weightedTotal = weightedTotal + CurrentValue
     Next i
'divide the WeightedTotal by 103 and get the remainder, this is the CheckDigitValue
     CheckDigitValue = (weightedTotal Mod 103)
'Now that we have the CheckDigitValue, find the corresponding ASCII character from the table
     If CheckDigitValue < 95 And CheckDigitValue > 0 Then C128_CheckDigit = Chr(CheckDigitValue + 32)
     If CheckDigitValue > 94 Then C128_CheckDigit = Chr(CheckDigitValue + 100)
     If CheckDigitValue = 0 Then C128_CheckDigit = Chr(194)
'Check for spaces or "00" and print ASCII 194 instead
'place changes in DataToPrint
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
          CurrentChar = Mid(DataToEncode, i, 1)
          If CurrentChar = " " Then CurrentChar = Chr(194)
          DataToPrint = DataToPrint & CurrentChar
     Next i
'Get Printable String
     Printable_string = C128Start & DataToPrint & C128_CheckDigit & C128_Stop & " "
'Return the PrintableString
     Code128b = Printable_string
End Function


Public Function Code128c(DataToEncode As String, ReturnType As Integer) As String
'
' Copyright © IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com
'
' You may use our source code in your applications only if you are using barcode fonts created by IDautomation.com, Inc.
' and you do not remove the copyright notices in the source code.
'
' The purpose of this code is to calculate the Code 128 barcode from character set C
'
' You MUST use the fully functional Code 128 (dated 12/2000 or later)
' font for this code to create and print a proper barcode
'
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
' Check to make sure data is numeric
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
     DataToEncode = OnlyCorrectData
'Check for an even number of digits, add 0 if not even
     If (Len(DataToEncode) Mod 2) = 1 Then DataToEncode = "0" & DataToEncode
'Assign start & stop codes
     StartCode = Chr(205)
     StopCode = Chr(206)
'<<<< Calculate Modulo 103 Check Digit and generate DataToPrint >>>>
'Set WeightedTotal to the Code 128 value of the start character
     weightedTotal = 105
     WeightValue = 1
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength Step 2
    'Get the value of each number pair
          CurrentValue = Mid(DataToEncode, i, 2)
    'get the DataToPrint
          If CurrentValue < 95 And CurrentValue > 0 Then DataToPrint = DataToPrint & Chr(CurrentValue + 32)
          If CurrentValue > 94 Then DataToPrint = DataToPrint & Chr(CurrentValue + 100)
          If CurrentValue = 0 Then DataToPrint = DataToPrint & Chr(194)
    'multiply by the weighting character
          CurrentValue = CurrentValue * WeightValue
    'add the values together to get the weighted total
          weightedTotal = weightedTotal + CurrentValue
          WeightValue = WeightValue + 1
     Next i
'divide the WeightedTotal by 103 and get the remainder, this is the CheckDigitValue
     CheckDigitValue = (weightedTotal Mod 103)
'Now that we have the CheckDigitValue, find the corresponding ASCII character from the table
     If CheckDigitValue < 95 And CheckDigitValue > 0 Then C128_CheckDigit = Chr(CheckDigitValue + 32)
     If CheckDigitValue > 94 Then C128_CheckDigit = Chr(CheckDigitValue + 100)
     If CheckDigitValue = 0 Then C128_CheckDigit = Chr(194)
'ReturnType 0 returns data formatted to the barcode font
     If ReturnType = 0 Then Code128c = StartCode & DataToPrint & C128_CheckDigit & StopCode & " "
'ReturnType 1 returns data formatted for human readable text
     If ReturnType = 1 Then Code128c = DataToEncode & CheckDigitValue
'ReturnType 2 returns the check digit for the data supplied
     If ReturnType = 2 Then Code128c = STR(CheckDigitValue)
End Function


Public Function I2of5(DataToEncode As String) As String
'
' Copyright © IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com
'
' You may use our source code in your applications only if you are using barcode fonts created by IDautomation.com, Inc.
' and you do not remove the copyright notices in the source code.
'
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
     DataToEncode = OnlyCorrectData
'Check for an even number of digits, add 0 if not even
     If (Len(DataToEncode) Mod 2) = 1 Then DataToEncode = "0" & DataToEncode
'Assign start and stop codes
     StartCode = Chr(203)
     StopCode = Chr(204)
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength Step 2
    'Get the value of each number pair
          CurrentCharNum = Val((Mid(DataToEncode, i, 2)))
    'Get the ASCII value of CurrentChar according to chart by to the value
          If CurrentCharNum < 94 Then DataToPrint = DataToPrint & Chr(CurrentCharNum + 33)
          If CurrentCharNum > 93 Then DataToPrint = DataToPrint & Chr(CurrentCharNum + 103)
     Next i
'Get Printable String
     Printable_string = StartCode + DataToPrint + StopCode & " "
'Return PrintableString
     I2of5 = Printable_string
End Function





Public Function USPSss(DataToEncode As String, ReturnType As Integer) As String
' This code is Copyright, IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com
' The purpose for this code is to print a Code 128 barcode
' according to the USPS standards.
     
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
     
'Remove check digits and (AI) if they were added to input
     If Len(OnlyCorrectData) = "20" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 19))
'End sub if incorrect number
     If Len(OnlyCorrectData) <> "19" Then End
     DataToEncode = OnlyCorrectData
     
'<<<< Generate MOD 10 check digit >>>>
     Factor = 3
     weightedTotal = 0
     StringLength = Len(DataToEncode)
     For i = StringLength To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, i, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          weightedTotal = weightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next i
'Find the CheckDigit by finding the smallest number that = a multiple of 10
     i = (weightedTotal Mod 10)
     If i <> 0 Then
          CheckDigit = (10 - i)
     Else
          CheckDigit = 0
     End If
'Now that we have calculated the MOD 10 for the data, send the string
'to the Code128c() funtion. This function will:
' - Add in the start and stop codes
' - Calculate the MOD 103 required by SS when using Code 128
' - Interleave the numbers into printable characters
'ReturnType 0 returns data formatted to the barcode font
     If ReturnType = 0 Then USPSss = Code128c(DataToEncode & CheckDigit, 0)
'ReturnType 1 returns data formatted for human readable text
     If ReturnType = 1 Then USPSss = Mid(DataToEncode, 1, 4) & " " & Mid(DataToEncode, 5, 4) & " " & Mid(DataToEncode, 9, 4) & " " & Mid(DataToEncode, 13, 4) & " " & Mid(DataToEncode, 17, 3) & CheckDigit
'ReturnType 2 returns the MOD10 check digit for the data supplied
     If ReturnType = 2 Then USPSss = STR(CheckDigit)
End Function




Public Function Code39Mod43(DataToEncode As String, ReturnType As Integer) As String
'
' This module is Copyright, IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com
'
' The purpose of this code is to print the MOD43 CODE 39 barcode
     DataToEncode = RTrim(DataToEncode)
     DataToEncode = UCase(DataToEncode)
     DataToPrint = ""
     OnlyCorrectData = ""
'only pass correct data
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Get each character one at a time
          CurrentCharNum = (Asc(Mid(DataToEncode, i, 1)))
    'Get the value of CurrentChar according to MOD43
    '0-9
          If CurrentCharNum < 58 And CurrentCharNum > 47 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
    'A-Z
          If CurrentCharNum < 91 And CurrentCharNum > 64 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
    'Space
          If CurrentCharNum = 32 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
    '-
          If CurrentCharNum = 45 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
    '.
          If CurrentCharNum = 46 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
    '$
          If CurrentCharNum = 36 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
    '/
          If CurrentCharNum = 47 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
    '+
          If CurrentCharNum = 43 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
    '%
          If CurrentCharNum = 37 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
     DataToEncode = OnlyCorrectData
     weightedTotal = 0
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Get each character one at a time
          CurrentCharNum = (Asc(Mid(DataToEncode, i, 1)))
    'Get the value of CurrentChar according to MOD43
    '0-9
          If CurrentCharNum < 58 And CurrentCharNum > 47 Then CurrentValue = CurrentCharNum - 48
    'A-Z
          If CurrentCharNum < 91 And CurrentCharNum > 64 Then CurrentValue = CurrentCharNum - 55
    'Space
          If CurrentCharNum = 32 Then CurrentValue = 38
    '-
          If CurrentCharNum = 45 Then CurrentValue = 36
    '.
          If CurrentCharNum = 46 Then CurrentValue = 37
    '$
          If CurrentCharNum = 36 Then CurrentValue = 39
    '/
          If CurrentCharNum = 47 Then CurrentValue = 40
    '+
          If CurrentCharNum = 43 Then CurrentValue = 41
    '%
          If CurrentCharNum = 37 Then CurrentValue = 42
    'To print the barcode symbol representing a space you will
    'to type or print "=" (the equal character) instead of a space character.
          If CurrentCharNum = 32 Then CurrentCharNum = 61
    'gather data to print
          DataToPrint = DataToPrint & Chr(CurrentCharNum)
    'add the values together
          weightedTotal = weightedTotal + CurrentValue
     Next i
'divide the WeightedTotal by 43 and get the remainder, this is the CheckDigit
     CheckDigitValue = (weightedTotal Mod 43)
    'Assign values to characters
    '0-9
     If CheckDigitValue < 10 Then CheckDigit = CheckDigitValue + 48
    'A-Z
     If CheckDigitValue < 36 And CheckDigitValue > 9 Then CheckDigit = CheckDigitValue + 55
    'Space
     If CheckDigitValue = 38 Then CheckDigit = 61
    '-
     If CheckDigitValue = 36 Then CheckDigit = 45
    '.
     If CheckDigitValue = 37 Then CheckDigit = 46
    '$
     If CheckDigitValue = 39 Then CheckDigit = 36
    '/
     If CheckDigitValue = 40 Then CheckDigit = 47
    '+
     If CheckDigitValue = 41 Then CheckDigit = 43
    '%
     If CheckDigitValue = 42 Then CheckDigit = 37
     
'ReturnType 0 returns data formatted to the barcode font
     If ReturnType = 0 Then Code39Mod43 = "!" & DataToPrint & Chr(CheckDigit) & "!" & " "
'ReturnType 1 returns data formatted for human readable text
     If ReturnType = 1 Then Code39Mod43 = DataToPrint & Chr(CheckDigit)
'ReturnType 2 returns the  check digit for the data supplied
     If ReturnType = 2 Then Code39Mod43 = Chr(CheckDigit)
End Function


Public Function Code39(DataToEncode As String) As String
'
' Copyright © IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com
'
' You may use our source code in your applications only if you are using barcode fonts created by IDautomation.com, Inc.
' and you do not remove the copyright notices in the source code.
'
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
'Check for spaces in code
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Get each character one at a time
          CurrentChar = (Mid(DataToEncode, i, 1))
    'To print the barcode symbol representing a space you will
    'to type or print "=" (the equal character) instead of a space character.
          If CurrentChar = " " Then CurrentChar = "="
          DataToPrint = DataToPrint & CurrentChar
     Next i
'Get Printable String
     Printable_string = "!" & DataToPrint & "!" & " "
'Return PrintableString
     Code39 = Printable_string
End Function





Public Function I2of5Mod10(DataToEncode As String, ReturnType As Integer) As String
'
' This module is Copyright, IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com or http://www.IDautomation.com
'
' The purpose of this code is to print the Interleaved 2 of 5 barcode
' With a MOD 10 check digit. This is now required by the US Post Office for
' printing barcodes on US MAIL for their "Special Services". Use the AdvI25p
' font when printing barcodes for US MAIL.
'
' Get data from user, this is the DataToEncode
     DataToEncode = RTrim(LTrim(DataToEncode))
     DataToPrint = ""
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
     DataToEncode = OnlyCorrectData
'<<<< Calculate Check Digit >>>>
     Factor = 3
     weightedTotal = 0
     For i = Len(DataToEncode) To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, i, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          weightedTotal = weightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next i
'Find the CheckDigit by finding the smallest number that = a multiple of 10
     i = (weightedTotal Mod 10)
     If i <> 0 Then
          CheckDigit = (10 - i)
     Else
          CheckDigit = 0
     End If
'Add check digit to number to DataToEncode
     DataToEncode = DataToEncode & CheckDigit
'Check for an even number of digits, add 0 if not even
     If (Len(DataToEncode) Mod 2) = 1 Then DataToEncode = "0" & DataToEncode
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength Step 2
    'Get the value of each number pair
          CurrentCharNum = (Mid(DataToEncode, i, 2))
    'Get the ASCII value of CurrentChar according to chart by to the value
          If CurrentCharNum < 94 Then DataToPrint = DataToPrint & Chr(CurrentCharNum + 33)
          If CurrentCharNum > 93 Then DataToPrint = DataToPrint & Chr(CurrentCharNum + 103)
     Next i
'ReturnType 0 returns data formatted to the barcode font
     If ReturnType = 0 Then I2of5Mod10 = Chr(203) & DataToPrint & Chr(204) & " "
'ReturnType 1 returns data formatted for human readable text
     If ReturnType = 1 Then I2of5Mod10 = DataToEncode
'ReturnType 2 returns the  check digit for the data supplied
     If ReturnType = 2 Then I2of5Mod10 = STR$(CheckDigit)
End Function





Public Function MSI(DataToEncode As String, ReturnType As Integer) As String
'
' Copyright © IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com or http://www.IDautomation.com
'
' You may use our source code in your applications only if you are using barcode fonts created by IDautomation.com, Inc.
' and you do not remove the copyright notices in the source code.
'
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
     DataToEncode = OnlyCorrectData
'<<<< Calculate Check Digit >>>>
Dim OddNumbers As String
Dim EvenNumberSum As Integer
Dim OddProductSum As Integer
Factor = 3
weightedTotal = "0"
StringLength = Len(DataToEncode)
For i = 1 To StringLength 'Step 1
    'Get the value of each EVEN number, 1st number is even & add them together
    If Factor = 1 Then EvenNumberSum = EvenNumberSum + Val(Mid(DataToEncode, i, 1))
    'Get the value of each ODD number, 2nd number is odd & gether them
    If Factor = 3 Then OddNumbers = OddNumbers & Val(Mid(DataToEncode, i, 1))
    Factor = 4 - Factor
Next i
'Multiply odd number gathered by 2
OddNumbers = OddNumbers * 2
'Add the digits of the product together
OddProductSum = "0"
For i = 1 To Len(OddNumbers)
    OddProductSum = OddProductSum + Val(Mid(OddNumbers, i, 1))
Next i
weightedTotal = OddProductSum + EvenNumberSum
'Find the CheckDigit by finding the number + weightedTotal that = a multiple of 10
'divide by 10, get the remainder and subtract from 10
     i = (weightedTotal Mod 10)
     If i <> 0 Then
          CheckDigit = (10 - i)
     Else
          CheckDigit = 0
     End If
'ReturnType 0 returns data formatted to the barcode font
     If ReturnType = 0 Then MSI = "(" & DataToEncode & CheckDigit & ")" & " "
'ReturnType 1 returns data formatted for human readable text
     If ReturnType = 1 Then MSI = DataToEncode
'ReturnType 2 returns the  check digit for the data supplied
     If ReturnType = 2 Then MSI = STR$(CheckDigit)
End Function


Public Function UPCa(DataToEncode As String) As String
'
' Copyright © IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com or http://www.IDautomation.com
'
' You may use our source code in your applications only if you are using barcode fonts created by IDautomation.com, Inc.
' and you do not remove the copyright notices in the source code.
'
' The purpose of this code is to calculate the UPC-A barcode
' Enter all the numbers without dashes
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
'Remove check digits if they added one
     If Len(OnlyCorrectData) = "12" Then OnlyCorrectData = Mid(OnlyCorrectData, 1, 11)
     If Len(OnlyCorrectData) = "14" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 11) & Mid(OnlyCorrectData, 13, 2))
     If Len(OnlyCorrectData) = "17" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 11) & Mid(OnlyCorrectData, 13, 5))
     EAN2AddOn = ""
     EAN5AddOn = ""
     EANAddOnToPrint = ""
     If Len(OnlyCorrectData) = 16 Then EAN5AddOn = Mid(OnlyCorrectData, 12, 5)
     If Len(OnlyCorrectData) = 13 Then EAN2AddOn = Mid(OnlyCorrectData, 12, 2)
'split 12 digit number from add-on
     DataToEncode = Mid(OnlyCorrectData, 1, 11)
'<<<< Calculate Check Digit >>>>
     Factor = 3
     weightedTotal = 0
     For i = Len(DataToEncode) To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, i, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          weightedTotal = weightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next i
'Find the CheckDigit by finding the number + weightedTotal that = a multiple of 10
'divide by 10, get the remainder and subtract from 10
     i = (weightedTotal Mod 10)
     If i <> 0 Then
          CheckDigit = (10 - i)
     Else
          CheckDigit = 0
     End If
     DataToEncode = DataToEncode & CheckDigit
'Now that have the total number including the check digit, determine character to print
'for proper barcoding
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Get the ASCII value of each number
          CurrentCharNum = Asc(Mid(DataToEncode, i, 1))
    'Print different barcodes according to the location of the CurrentChar
          Select Case i
          Case 1
        'For the first character print the human readable character, the normal
        'guard pattern and then the barcode without the human readable character
               If Chr(CurrentCharNum) > 4 Then DataToPrint = Chr(CurrentCharNum + 64) & "(" & Chr(CurrentCharNum + 49)
               If Chr(CurrentCharNum) < 5 Then DataToPrint = Chr(CurrentCharNum + 37) & "(" & Chr(CurrentCharNum + 49)
          Case 2
               DataToPrint = DataToPrint & Chr(CurrentCharNum)
          Case 3
               DataToPrint = DataToPrint & Chr(CurrentCharNum)
          Case 4
               DataToPrint = DataToPrint & Chr(CurrentCharNum)
          Case 5
               DataToPrint = DataToPrint & Chr(CurrentCharNum)
          Case 6
        'Print the center guard pattern after the 6th character
               DataToPrint = DataToPrint & Chr(CurrentCharNum) & "*"
          Case 7
        'Add 27 to the ASII value of characters 6-12 to print from character set+ C
        'this is required when printing to the right of the center guard pattern
               DataToPrint = DataToPrint & Chr(CurrentCharNum + 27)
          Case 8
               DataToPrint = DataToPrint & Chr(CurrentCharNum + 27)
          Case 9
               DataToPrint = DataToPrint & Chr(CurrentCharNum + 27)
          Case 10
               DataToPrint = DataToPrint & Chr(CurrentCharNum + 27)
          Case 11
               DataToPrint = DataToPrint & Chr(CurrentCharNum + 27)
          Case 12
        'For the last character print the barcode without the human readable character,
        'the normal guard pattern and then the human readable character.
               If Chr(CurrentCharNum) > 4 Then DataToPrint = DataToPrint & Chr(CurrentCharNum + 59) & "(" & Chr(CurrentCharNum + 64)
               If Chr(CurrentCharNum) < 5 Then DataToPrint = DataToPrint & Chr(CurrentCharNum + 59) & "(" & Chr(CurrentCharNum + 37)
          End Select
     Next i
'Process 5 digit add on if it exists
     If Len(EAN5AddOn) = 5 Then
          EANAddOnToPrint = ""
    'Get check digit for add on
          Factor = 3
          weightedTotal = 0
          For i = Len(EAN5AddOn) To 1 Step -1
        'Get the value of each number starting at the end
               CurrentCharNum = Mid(EAN5AddOn, i, 1)
        'multiply by the weighting factor which is 3,9,3,9.
        'and add the sum together
               If Factor = 3 Then weightedTotal = weightedTotal + CurrentCharNum * 3
               If Factor = 1 Then weightedTotal = weightedTotal + CurrentCharNum * 9
        'change factor for next calculation
               Factor = 4 - Factor
          Next i
    'Find the CheckDigit by extracting the right-most number from weightedTotal
          CheckDigit = Val(Right$(weightedTotal, 1))
    'Now we must encode the add-on CheckDigit into the number sets
    'by using variable parity between character sets A and B
          Select Case CheckDigit
          Case 0
               Encoding = "BBAAA"
          Case 1
               Encoding = "BABAA"
          Case 2
               Encoding = "BAABA"
          Case 3
               Encoding = "BAAAB"
          Case 4
               Encoding = "ABBAA"
          Case 5
               Encoding = "AABBA"
          Case 6
               Encoding = "AAABB"
          Case 7
               Encoding = "ABABA"
          Case 8
               Encoding = "ABAAB"
          Case 9
               Encoding = "AABAB"
          End Select
    'Now that we have the total number including the check digit, determine character to print
    'for proper barcoding:
          For i = 1 To Len(EAN5AddOn)
        'Get the value of each number
        'it is encoded with variable parity
               CurrentChar = Mid(EAN5AddOn, i, 1)
               CurrentEncoding = Mid(Encoding, i, 1)
        'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
               Select Case CurrentEncoding
               Case "A"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & Chr(34)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & Chr(35)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & Chr(36)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & Chr(37)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & Chr(38)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & Chr(44)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & Chr(46)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & Chr(47)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & Chr(58)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & Chr(59)
               Case "B"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & Chr(122)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & Chr(61)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & Chr(63)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & Chr(64)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & Chr(91)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & Chr(92)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & Chr(93)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & Chr(95)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & Chr(123)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & Chr(125)
               End Select
        'add in the space & add-on guard pattern
               Select Case i
               Case 1
            'EANAddOnToPrint = Chr(32) & Chr(43) & EANAddOnToPrint & Chr(33)
                    EANAddOnToPrint = Chr(43) & EANAddOnToPrint & Chr(33)
            'Now print add-on delineators between each add-on character
               Case 2
                    EANAddOnToPrint = EANAddOnToPrint & Chr(33)
               Case 3
                    EANAddOnToPrint = EANAddOnToPrint & Chr(33)
               Case 4
                    EANAddOnToPrint = EANAddOnToPrint & Chr(33)
               Case 5
                    EANAddOnToPrint = EANAddOnToPrint
               End Select
          Next i
     End If
'Process 2 digit add on if it exists
     If Len(EAN2AddOn) = 2 Then
          EANAddOnToPrint = ""
    'Get encoding for add on
          For i = 0 To 99 Step 4
               If Val(EAN2AddOn) = i Then Encoding = "AA"
               If Val(EAN2AddOn) = i + 1 Then Encoding = "AB"
               If Val(EAN2AddOn) = i + 2 Then Encoding = "BA"
               If Val(EAN2AddOn) = i + 3 Then Encoding = "BB"
          Next i
    'Now that we have the total number including the encoding
    'determine what to print
          For i = 1 To Len(EAN2AddOn)
        'Get the value of each number
        'it is encoded with variable parity
               CurrentChar = Mid(EAN2AddOn, i, 1)
               CurrentEncoding = Mid(Encoding, i, 1)
        'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
               Select Case CurrentEncoding
               Case "A"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & Chr(34)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & Chr(35)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & Chr(36)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & Chr(37)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & Chr(38)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & Chr(44)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & Chr(46)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & Chr(47)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & Chr(58)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & Chr(59)
               Case "B"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & Chr(122)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & Chr(61)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & Chr(63)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & Chr(64)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & Chr(91)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & Chr(92)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & Chr(93)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & Chr(95)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & Chr(123)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & Chr(125)
               End Select
        'add in the space & add-on guard pattern
               Select Case i
               Case 1
            'EANAddOnToPrint = Chr(32) & Chr(43) & EANAddOnToPrint & Chr(33)
                    EANAddOnToPrint = Chr(43) & EANAddOnToPrint & Chr(33)
            'Now print add-on delineators between each add-on character
               Case 2
                    EANAddOnToPrint = EANAddOnToPrint
               End Select
          Next i
     End If
'Get Printable String
     Printable_string = DataToPrint & EANAddOnToPrint & " "
'Return PrintableString
     UPCa = Printable_string
End Function



Public Function UPCe(DataToEncode As String) As String
'
' This module is Copyright, IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com or http://www.IDautomation.com
'
' The purpose of this code is to print the UPC-E barcode
' from a UPC-A barcode that can be compressed.
'
' Get data from user, this is the DataToEncode
     DataToEncode = RTrim(LTrim(DataToEncode))
     DataToPrint = ""
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
'Remove check digits if they added one
     If Len(OnlyCorrectData) = "12" Then OnlyCorrectData = Mid(OnlyCorrectData, 1, 11)
     If Len(OnlyCorrectData) = "14" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 11) & Mid(OnlyCorrectData, 13, 2))
     If Len(OnlyCorrectData) = "17" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 11) & Mid(OnlyCorrectData, 13, 5))
     EAN2AddOn = ""
     EAN5AddOn = ""
     EANAddOnToPrint = ""
     If Len(OnlyCorrectData) = 16 Then EAN5AddOn = Mid(OnlyCorrectData, 12, 5)
     If Len(OnlyCorrectData) = 13 Then EAN2AddOn = Mid(OnlyCorrectData, 12, 2)
'split 12 digit number from add-on
     DataToEncode = Mid(OnlyCorrectData, 1, 11)
     
'<<<< Calculate Check Digit >>>>
     Factor = 3
     weightedTotal = 0
     For i = Len(DataToEncode) To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, i, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          weightedTotal = weightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next i
'Find the CheckDigit by finding the number + weightedTotal that = a multiple of 10
'divide by 10, get the remainder and subtract from 10
     i = (weightedTotal Mod 10)
     If i <> 0 Then
          CheckDigit = (10 - i)
     Else
          CheckDigit = 0
     End If
     
     DataToEncode = DataToEncode & CheckDigit
'Compress UPC-A to UPC-E if possible
     Dim D1 As String
     Dim D2 As String
     Dim D3 As String
     Dim D4 As String
     Dim D5 As String
     Dim D6 As String
     Dim D7 As String
     Dim D8 As String
     Dim D9 As String
     Dim D10 As String
     Dim D11 As String
     Dim D12 As String
     D1 = Mid(DataToEncode, 1, 1)
     D2 = Mid(DataToEncode, 2, 1)
     D3 = Mid(DataToEncode, 3, 1)
     D4 = Mid(DataToEncode, 4, 1)
     D5 = Mid(DataToEncode, 5, 1)
     D6 = Mid(DataToEncode, 6, 1)
     D7 = Mid(DataToEncode, 7, 1)
     D8 = Mid(DataToEncode, 8, 1)
     D9 = Mid(DataToEncode, 9, 1)
     D10 = Mid(DataToEncode, 10, 1)
     D11 = Mid(DataToEncode, 11, 1)
     D12 = Mid(DataToEncode, 12, 1)
'Condition A
     If (D11 = "5" Or D11 = "6" Or D11 = "7" Or D11 = "8" Or D11 = "9") And D6 <> "0" And (D7 = "0" And D8 = "0" And D9 = "0" And D10 = "0") Then
          DataToEncode = D2 & D3 & D4 & D5 & D6 & D11
     End If
'Condition B
     If (D6 = "0" And D7 = "0" And D8 = "0" And D9 = "0" And D10 = "0") And D5 <> "0" Then
          DataToEncode = D2 & D3 & D4 & D5 & D11 & "4"
     End If
'Condition C
     If (D5 = "0" And D6 = "0" And D7 = "0" And D8 = "0") And (D4 = "1" Or D4 = "2" Or D4 = "0") Then
          DataToEncode = D2 & D3 & D9 & D10 & D11 & D4
     End If
'Condition D
     If (D5 = "0" And D6 = "0" And D7 = "0" And D8 = "0" And D9 = "0") And (D4 = "3" Or D4 = "4" Or D4 = "5" Or D4 = "6" Or D4 = "7" Or D4 = "8" Or D4 = "9") Then
          DataToEncode = D2 & D3 & D4 & D10 & D11 & "3"
     End If
'
'Run UPC-E compression only if DataToEncode = 6
     If Len(DataToEncode) = 6 Then
    'Now we must encode the check character into the symbol
    'by using variable parity between character sets A and B
          Select Case D12
          Case "0"
               Encoding = "BBBAAA"
          Case "1"
               Encoding = "BBABAA"
          Case "2"
               Encoding = "BBAABA"
          Case "3"
               Encoding = "BBAAAB"
          Case "4"
               Encoding = "BABBAA"
          Case "5"
               Encoding = "BAABBA"
          Case "6"
               Encoding = "BAAABB"
          Case "7"
               Encoding = "BABABA"
          Case "8"
               Encoding = "BABAAB"
          Case "9"
               Encoding = "BAABAB"
          End Select
          StringLength = Len(DataToEncode)
          For i = 1 To StringLength
        'Get the ASCII value of each number
               CurrentCharNum = Asc(Mid(DataToEncode, i, 1))
               CurrentEncoding = Mid(Encoding, i, 1)
        'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
               Select Case CurrentEncoding
               Case "A"
                    DataToPrint = DataToPrint & Chr(CurrentCharNum)
               Case "B"
                    DataToPrint = DataToPrint & Chr(CurrentCharNum + 17)
               End Select
        'add in the 1st character along with guard patterns
               Select Case i
               Case 1
            'For the LeadingDigit print the human readable character,
            'the normal guard pattern and then the rest of the barcode
                    DataToPrint = Chr(85) & "(" & DataToPrint
               Case 6
            'Print the SPECIAL guard pattern and check character
                    If CInt(D12) > 4 Then DataToPrint = DataToPrint & ")" & Chr(Asc(D12) + 64)
                    If CInt(D12) < 5 Then DataToPrint = DataToPrint & ")" & Chr(Asc(D12) + 37)
                    
               End Select
          Next i
     End If
     
'determine character to print
'for proper upc-a barcoding
     If Len(DataToEncode) <> 6 Then
          StringLength = Len(DataToEncode)
          For i = 1 To StringLength
        'Get the ASCII value of each number
               CurrentCharNum = Asc(Mid(DataToEncode, i, 1))
        'Print different barcodes according to the location of the CurrentChar
               Select Case i
               Case 1
            'For the first character print the human readable character, the normal
            'guard pattern and then the barcode without the human readable character
                    If Chr(CurrentCharNum) > 4 Then DataToPrint = Chr(CurrentCharNum + 64) & "(" & Chr(CurrentCharNum + 49)
                    If Chr(CurrentCharNum) < 5 Then DataToPrint = Chr(CurrentCharNum + 37) & "(" & Chr(CurrentCharNum + 49)
               Case 2
                    DataToPrint = DataToPrint & Chr(CurrentCharNum)
               Case 3
                    DataToPrint = DataToPrint & Chr(CurrentCharNum)
               Case 4
                    DataToPrint = DataToPrint & Chr(CurrentCharNum)
               Case 5
                    DataToPrint = DataToPrint & Chr(CurrentCharNum)
               Case 6
            'Print the center guard pattern after the 6th character
                    DataToPrint = DataToPrint & Chr(CurrentCharNum) & "*"
               Case 7
            'Add 27 to the ASII value of characters 6-12 to print from character set+ C
            'this is required when printing to the right of the center guard pattern
                    DataToPrint = DataToPrint & Chr(CurrentCharNum + 27)
               Case 8
                    DataToPrint = DataToPrint & Chr(CurrentCharNum + 27)
               Case 9
                    DataToPrint = DataToPrint & Chr(CurrentCharNum + 27)
               Case 10
                    DataToPrint = DataToPrint & Chr(CurrentCharNum + 27)
               Case 11
                    DataToPrint = DataToPrint & Chr(CurrentCharNum + 27)
               Case 12
            'For the last character print the barcode without the human readable character,
            'the normal guard pattern and then the human readable character.
                    If Chr(CurrentCharNum) > 4 Then DataToPrint = DataToPrint & Chr(CurrentCharNum + 59) & "(" & Chr(CurrentCharNum + 64)
                    If Chr(CurrentCharNum) < 5 Then DataToPrint = DataToPrint & Chr(CurrentCharNum + 59) & "(" & Chr(CurrentCharNum + 37)
               End Select
          Next i
     End If
     
'Process 5 digit add on if it exists
     If Len(EAN5AddOn) = 5 Then
          EANAddOnToPrint = ""
    'Get check digit for add on
          Factor = 3
          weightedTotal = 0
          For i = Len(EAN5AddOn) To 1 Step -1
        'Get the value of each number starting at the end
               CurrentCharNum = Mid(EAN5AddOn, i, 1)
        'multiply by the weighting factor which is 3,9,3,9.
        'and add the sum together
               If Factor = 3 Then weightedTotal = weightedTotal + CurrentCharNum * 3
               If Factor = 1 Then weightedTotal = weightedTotal + CurrentCharNum * 9
        'change factor for next calculation
               Factor = 4 - Factor
          Next i
    'Find the CheckDigit by extracting the right-most number from weightedTotal
          CheckDigit = Val(Right$(weightedTotal, 1))
    'Now we must encode the add-on CheckDigit into the number sets
    'by using variable parity between character sets A and B
          Select Case CheckDigit
          Case 0
               Encoding = "BBAAA"
          Case 1
               Encoding = "BABAA"
          Case 2
               Encoding = "BAABA"
          Case 3
               Encoding = "BAAAB"
          Case 4
               Encoding = "ABBAA"
          Case 5
               Encoding = "AABBA"
          Case 6
               Encoding = "AAABB"
          Case 7
               Encoding = "ABABA"
          Case 8
               Encoding = "ABAAB"
          Case 9
               Encoding = "AABAB"
          End Select
          
    'Now that we have the total number including the check digit, determine character to print
    'for proper barcoding:
          For i = 1 To Len(EAN5AddOn)
        'Get the value of each number
        'it is encoded with variable parity
               CurrentChar = Mid(EAN5AddOn, i, 1)
               CurrentEncoding = Mid(Encoding, i, 1)
        'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
               Select Case CurrentEncoding
               Case "A"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & Chr(34)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & Chr(35)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & Chr(36)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & Chr(37)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & Chr(38)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & Chr(44)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & Chr(46)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & Chr(47)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & Chr(58)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & Chr(59)
               Case "B"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & Chr(122)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & Chr(61)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & Chr(63)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & Chr(64)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & Chr(91)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & Chr(92)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & Chr(93)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & Chr(95)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & Chr(123)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & Chr(125)
               End Select
        'add in the space & add-on guard pattern
               Select Case i
               Case 1
            'EANAddOnToPrint = Chr(32) & Chr(43) & EANAddOnToPrint & Chr(33)
                    EANAddOnToPrint = Chr(43) & EANAddOnToPrint & Chr(33)
            'Now print add-on delineators between each add-on character
               Case 2
                    EANAddOnToPrint = EANAddOnToPrint & Chr(33)
               Case 3
                    EANAddOnToPrint = EANAddOnToPrint & Chr(33)
               Case 4
                    EANAddOnToPrint = EANAddOnToPrint & Chr(33)
               Case 5
                    EANAddOnToPrint = EANAddOnToPrint
               End Select
          Next i
     End If
     
'Process 2 digit add on if it exists
     If Len(EAN2AddOn) = 2 Then
          EANAddOnToPrint = ""
    'Get encoding for add on
          For i = 0 To 99 Step 4
               If Val(EAN2AddOn) = i Then Encoding = "AA"
               If Val(EAN2AddOn) = i + 1 Then Encoding = "AB"
               If Val(EAN2AddOn) = i + 2 Then Encoding = "BA"
               If Val(EAN2AddOn) = i + 3 Then Encoding = "BB"
          Next i
    'Now that we have the total number including the encoding
    'determine what to print
          For i = 1 To Len(EAN2AddOn)
        'Get the value of each number
        'it is encoded with variable parity
               CurrentChar = Mid(EAN2AddOn, i, 1)
               CurrentEncoding = Mid(Encoding, i, 1)
        'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
               Select Case CurrentEncoding
               Case "A"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & Chr(34)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & Chr(35)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & Chr(36)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & Chr(37)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & Chr(38)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & Chr(44)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & Chr(46)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & Chr(47)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & Chr(58)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & Chr(59)
               Case "B"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & Chr(122)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & Chr(61)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & Chr(63)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & Chr(64)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & Chr(91)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & Chr(92)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & Chr(93)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & Chr(95)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & Chr(123)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & Chr(125)
               End Select
        'add in the space & add-on guard pattern
               Select Case i
               Case 1
            'EANAddOnToPrint = Chr(32) & Chr(43) & EANAddOnToPrint & Chr(33)
                    EANAddOnToPrint = Chr(43) & EANAddOnToPrint & Chr(33)
            'Now print add-on delineators between each add-on character
               Case 2
                    EANAddOnToPrint = EANAddOnToPrint
               End Select
          Next i
     End If
     
'Get Printable String
     Printable_string = DataToPrint & EANAddOnToPrint & " "
     
'Return PrintableString
     UPCe = Printable_string
     
End Function



Public Function EAN13(DataToEncode As String) As String
'
' Copyright © IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com or http://www.IDautomation.com
'
' You may use our source code in your applications only if you are using barcode fonts created by IDautomation.com, Inc.
' and you do not remove the copyright notices in the source code.
'
' The purpose of this code is to calculate the EAN-13 barcode
'
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
'DataToEncode = OnlyCorrectData
''
'Remove check digits if they added one
     If Len(OnlyCorrectData) = "13" Then OnlyCorrectData = Mid(OnlyCorrectData, 1, 12)
     If Len(OnlyCorrectData) = "15" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 12) & Mid(OnlyCorrectData, 14, 2))
     If Len(OnlyCorrectData) = "18" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 12) & Mid(OnlyCorrectData, 14, 5))
'End sub if incorrect number
     Dim EAN2AddOn As String
     Dim EAN5AddOn As String
     Dim EANAddOnToPrint As String
     EAN2AddOn = ""
     EAN5AddOn = ""
     EANAddOnToPrint = ""
     If Len(OnlyCorrectData) = 17 Then EAN5AddOn = Mid(OnlyCorrectData, 13, 5)
     If Len(OnlyCorrectData) = 14 Then EAN2AddOn = Mid(OnlyCorrectData, 13, 2)
'split 12 digit number from add-on
     DataToEncode = Mid(OnlyCorrectData, 1, 12)
'<<<< Calculate Check Digit >>>>
     Factor = 3
     weightedTotal = 0
     For i = Len(DataToEncode) To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, i, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          weightedTotal = weightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next i
'Find the CheckDigit by finding the number + weightedTotal that = a multiple of 10
'divide by 10, get the remainder and subtract from 10
     i = (weightedTotal Mod 10)
     If i <> 0 Then
          CheckDigit = (10 - i)
     Else
          CheckDigit = 0
     End If
'Now we must encode the leading digit into the left half of the EAN-13 symbol
'by using variable parity between character sets A and B
     LeadingDigit = Mid(DataToEncode, 1, 1)
     Select Case LeadingDigit
     Case 0
          Encoding = "AAAAAACCCCCC"
     Case 1
          Encoding = "AABABBCCCCCC"
     Case 2
          Encoding = "AABBABCCCCCC"
     Case 3
          Encoding = "AABBBACCCCCC"
     Case 4
          Encoding = "ABAABBCCCCCC"
     Case 5
          Encoding = "ABBAABCCCCCC"
     Case 6
          Encoding = "ABBBAACCCCCC"
     Case 7
          Encoding = "ABABABCCCCCC"
     Case 8
          Encoding = "ABABBACCCCCC"
     Case 9
          Encoding = "ABBABACCCCCC"
     End Select
'add the check digit to the end of the barcode & remove the leading digit
     DataToEncode = Mid(DataToEncode, 2, 11) & CheckDigit
'Now that we have the total number including the check digit, determine character to print
'for proper barcoding:
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Get the ASCII value of each number excluding the first number because
    'it is encoded with variable parity
          CurrentCharNum = Asc(Mid(DataToEncode, i, 1))
          CurrentEncoding = Mid(Encoding, i, 1)
    'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
          Select Case CurrentEncoding
          Case "A"
               DataToPrint = DataToPrint & Chr(CurrentCharNum)
          Case "B"
               DataToPrint = DataToPrint & Chr(CurrentCharNum + 17)
          Case "C"
               DataToPrint = DataToPrint & Chr(CurrentCharNum + 27)
          End Select
    'add in the 1st character along with guard patterns
          Select Case i
          Case 1
        'For the LeadingDigit print the human readable character,
        'the normal guard pattern and then the rest of the barcode
               If LeadingDigit > 4 Then DataToPrint = Chr((LeadingDigit + 48) + 64) & "(" & DataToPrint
               If LeadingDigit < 5 Then DataToPrint = Chr((LeadingDigit + 48) + 37) & "(" & DataToPrint
          Case 6
        'Print the center guard pattern after the 6th character
               DataToPrint = DataToPrint & "*"
          Case 12
        'For the last character (12) print the the normal guard pattern
        'after the barcode
               DataToPrint = DataToPrint & "("
          End Select
     Next i
'Process 5 digit add on if it exists
     If Len(EAN5AddOn) = 5 Then
          EANAddOnToPrint = ""
    'Get check digit for add on
          Factor = 3
          weightedTotal = 0
          For i = Len(EAN5AddOn) To 1 Step -1
        'Get the value of each number starting at the end
               CurrentCharNum = Mid(EAN5AddOn, i, 1)
        'multiply by the weighting factor which is 3,9,3,9.
        'and add the sum together
               If Factor = 3 Then weightedTotal = weightedTotal + CurrentCharNum * 3
               If Factor = 1 Then weightedTotal = weightedTotal + CurrentCharNum * 9
        'change factor for next calculation
               Factor = 4 - Factor
          Next i
    'Find the CheckDigit by extracting the right-most number from weightedTotal
          CheckDigit = Val(Right$(weightedTotal, 1))
    'Now we must encode the add-on CheckDigit into the number sets
    'by using variable parity between character sets A and B
          Select Case CheckDigit
          Case 0
               Encoding = "BBAAA"
          Case 1
               Encoding = "BABAA"
          Case 2
               Encoding = "BAABA"
          Case 3
               Encoding = "BAAAB"
          Case 4
               Encoding = "ABBAA"
          Case 5
               Encoding = "AABBA"
          Case 6
               Encoding = "AAABB"
          Case 7
               Encoding = "ABABA"
          Case 8
               Encoding = "ABAAB"
          Case 9
               Encoding = "AABAB"
          End Select
    'Now that we have the total number including the check digit, determine character to print
    'for proper barcoding:
          For i = 1 To Len(EAN5AddOn)
        'Get the value of each number
        'it is encoded with variable parity
               CurrentChar = Mid(EAN5AddOn, i, 1)
               CurrentEncoding = Mid(Encoding, i, 1)
        'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
               Select Case CurrentEncoding
               Case "A"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & Chr(34)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & Chr(35)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & Chr(36)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & Chr(37)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & Chr(38)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & Chr(44)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & Chr(46)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & Chr(47)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & Chr(58)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & Chr(59)
               Case "B"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & Chr(122)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & Chr(61)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & Chr(63)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & Chr(64)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & Chr(91)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & Chr(92)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & Chr(93)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & Chr(95)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & Chr(123)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & Chr(125)
               End Select
        'add in the space & add-on guard pattern
               Select Case i
               Case 1
                    EANAddOnToPrint = Chr(32) & Chr(43) & EANAddOnToPrint & Chr(33)
          'Now print add-on delineators between each add-on character
               Case 2
                    EANAddOnToPrint = EANAddOnToPrint & Chr(33)
               Case 3
                    EANAddOnToPrint = EANAddOnToPrint & Chr(33)
               Case 4
                    EANAddOnToPrint = EANAddOnToPrint & Chr(33)
               Case 5
                    EANAddOnToPrint = EANAddOnToPrint
               End Select
          Next i
     End If
'Process 2 digit add on if it exists
     If Len(EAN2AddOn) = 2 Then
          EANAddOnToPrint = ""
    'Get encoding for add on
          For i = 0 To 99 Step 4
               If Val(EAN2AddOn) = i Then Encoding = "AA"
               If Val(EAN2AddOn) = i + 1 Then Encoding = "AB"
               If Val(EAN2AddOn) = i + 2 Then Encoding = "BA"
               If Val(EAN2AddOn) = i + 3 Then Encoding = "BB"
          Next i
    'Now that we have the total number including the encoding
    'determine what to print
          For i = 1 To Len(EAN2AddOn)
        'Get the value of each number
        'it is encoded with variable parity
               CurrentChar = Mid(EAN2AddOn, i, 1)
               CurrentEncoding = Mid(Encoding, i, 1)
        'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
               Select Case CurrentEncoding
               Case "A"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & Chr(34)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & Chr(35)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & Chr(36)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & Chr(37)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & Chr(38)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & Chr(44)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & Chr(46)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & Chr(47)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & Chr(58)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & Chr(59)
               Case "B"
                    If CurrentChar = "0" Then EANAddOnToPrint = EANAddOnToPrint & Chr(122)
                    If CurrentChar = "1" Then EANAddOnToPrint = EANAddOnToPrint & Chr(61)
                    If CurrentChar = "2" Then EANAddOnToPrint = EANAddOnToPrint & Chr(63)
                    If CurrentChar = "3" Then EANAddOnToPrint = EANAddOnToPrint & Chr(64)
                    If CurrentChar = "4" Then EANAddOnToPrint = EANAddOnToPrint & Chr(91)
                    If CurrentChar = "5" Then EANAddOnToPrint = EANAddOnToPrint & Chr(92)
                    If CurrentChar = "6" Then EANAddOnToPrint = EANAddOnToPrint & Chr(93)
                    If CurrentChar = "7" Then EANAddOnToPrint = EANAddOnToPrint & Chr(95)
                    If CurrentChar = "8" Then EANAddOnToPrint = EANAddOnToPrint & Chr(123)
                    If CurrentChar = "9" Then EANAddOnToPrint = EANAddOnToPrint & Chr(125)
               End Select
        'add in the space & add-on guard pattern
               Select Case i
               Case 1
                    EANAddOnToPrint = Chr(32) & Chr(43) & EANAddOnToPrint & Chr(33)
          'Now print add-on delineators between each add-on character
               Case 2
                    EANAddOnToPrint = EANAddOnToPrint
               End Select
          Next i
     End If
'Get Printable String
     Printable_string = DataToPrint & EANAddOnToPrint & " "
'Return PrintableString
     EAN13 = Printable_string
End Function


Public Function EAN8(DataToEncode As String) As String
'
' Copyright © IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com or http://www.IDautomation.com
'
' You may use our source code in your applications only if you are using barcode fonts created by IDautomation.com, Inc.
' and you do not remove the copyright notices in the source code.
'
' The purpose of this code is to calculate the EAN-8 barcode
' Enter all the numbers without dashes
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
     DataToEncode = OnlyCorrectData
     If Len(DataToEncode) <> "7" Then
          MsgBox "Cannot process; you MUST enter a 7 digit NUMBER for this type of barcode. Do not use any spaces or dashes."
          Exit Function
     End If
'<<<< Calculate Check Digit >>>>
     Factor = 3
     weightedTotal = 0
     For i = Len(DataToEncode) To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, i, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          weightedTotal = weightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next i
'Find the CheckDigit by finding the number + weightedTotal that = a multiple of 10
'divide by 10, get the remainder and subtract from 10
     i = (weightedTotal Mod 10)
     If i <> 0 Then
          CheckDigit = (10 - i)
     Else
          CheckDigit = 0
     End If
     DataToEncode = DataToEncode & CheckDigit
'Now that have the total number including the check digit, determine character to print
'for proper barcoding
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Get the ASCII value of each number
          CurrentCharNum = Asc(Mid(DataToEncode, i, 1))
          CurrentEncoding = Mid(Encoding, i, 1)
    'Print different barcodes according to the location of the CurrentChar and CurrentEncoding
    'Print different barcodes according to the location of the CurrentChar
          Select Case i
          Case 1
        'For the first character print the normal guard pattern
        'and then the barcode without the human readable character
               DataToPrint = "(" & Chr(CurrentCharNum)
          Case 2
               DataToPrint = DataToPrint & Chr(CurrentCharNum)
          Case 3
               DataToPrint = DataToPrint & Chr(CurrentCharNum)
          Case 4
        'Print the center guard pattern after the 6th character
               DataToPrint = DataToPrint & Chr(CurrentCharNum) & "*"
          Case 5
               DataToPrint = DataToPrint & Chr(CurrentCharNum + 27)
          Case 6
               DataToPrint = DataToPrint & Chr(CurrentCharNum + 27)
          Case 7
               DataToPrint = DataToPrint & Chr(CurrentCharNum + 27)
          Case 8
        'Print the check digit as 8th character + normal guard pattern
               DataToPrint = DataToPrint & Chr(CurrentCharNum + 27) & "("
          End Select
     Next i
'Get Printable String
     Printable_string = DataToPrint & " "
'Display PrintableString in textbox
     EAN8 = Printable_string
End Function





Public Function SSCC18(DataToEncode As String, ReturnType As Integer) As String
' This code is Copyright, IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com
' The purpose for this code is to print a barcode
' according to the UCC/EAN SSCC-18 standards.
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
'Remove check digits and (AI) if they were added to input
     If Len(OnlyCorrectData) = "18" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 17))
     If Len(OnlyCorrectData) = "19" Then OnlyCorrectData = (Mid(OnlyCorrectData, 3, 17))
     If Len(OnlyCorrectData) = "20" Then OnlyCorrectData = (Mid(OnlyCorrectData, 3, 17))
     If Len(OnlyCorrectData) = "21" Then OnlyCorrectData = (Mid(OnlyCorrectData, 3, 17))
'End sub if incorrect number
     If Len(OnlyCorrectData) <> "17" Then End
     DataToEncode = OnlyCorrectData
'<<<< Generate MOD 10 check digit >>>>
     Factor = 3
     weightedTotal = 0
     StringLength = Len(DataToEncode)
     For i = StringLength To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, i, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          weightedTotal = weightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next i
'Find the CheckDigit by finding the smallest number that = a multiple of 10
     i = (weightedTotal Mod 10)
     If i <> 0 Then
          CheckDigit = (10 - i)
     Else
          CheckDigit = 0
     End If
'Add check digit and Application Identifier (AI) to DataToEncode
'AI = 00 for SSCC18
'DataToEncode = "00" & DataToEncode & CheckDigit
'Now that we have calculated the MOD 10 for the data, send the string
'to the UCC128() funtion. This function will:
' - Add in the Start C and FNC1 required by UCC/EAN
' - Calculate the MOD 103 required by UCC/EAN
' - Interleave the numbers into printable characters
'ReturnType 0 returns data formatted to the barcode font
     If ReturnType = 0 Then SSCC18 = UCC128("00" & DataToEncode & CheckDigit)
'ReturnType 1 returns data formatted for human readable text
     If ReturnType = 1 Then SSCC18 = "(00) " & Mid(DataToEncode, 1, 1) & " " & Mid(DataToEncode, 2, 7) & " " & Mid(DataToEncode, 9, 9) & " " & CheckDigit
'ReturnType 2 returns the MOD10 check digit for the data supplied
     If ReturnType = 2 Then SSCC18 = STR(CheckDigit)
End Function


Public Function SCC14(DataToEncode As String, ReturnType As Integer) As String
' This code is Copyright, IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com
' The purpose for this code is to print a barcode
' according to the UCC/EAN SCC-14 standards.
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
'Remove check digits and (AI) if they were added to input
     If Len(OnlyCorrectData) = "14" Then OnlyCorrectData = (Mid(OnlyCorrectData, 1, 13))
     If Len(OnlyCorrectData) = "15" Then OnlyCorrectData = (Mid(OnlyCorrectData, 3, 13))
     If Len(OnlyCorrectData) = "16" Then OnlyCorrectData = (Mid(OnlyCorrectData, 3, 13))
     If Len(OnlyCorrectData) = "17" Then OnlyCorrectData = (Mid(OnlyCorrectData, 3, 13))
'End sub if incorrect number
     If Len(OnlyCorrectData) <> "13" Then End
     DataToEncode = OnlyCorrectData
'<<<< Generate MOD 10 check digit >>>>
     Factor = 3
     weightedTotal = 0
     StringLength = Len(DataToEncode)
     For i = StringLength To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, i, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          weightedTotal = weightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next i
'Find the CheckDigit by finding the smallest number that = a multiple of 10
     i = (weightedTotal Mod 10)
     If i <> 0 Then
          CheckDigit = (10 - i)
     Else
          CheckDigit = 0
     End If
'Add check digit and Application Identifier (AI) to DataToEncode
'AI = 00 for SSCC18
'DataToEncode = "00" & DataToEncode & CheckDigit
'Now that we have calculated the MOD 10 for the data, send the string
'to the UCC128() funtion. This function will:
' - Add in the Start C and FNC1 required by UCC/EAN
' - Calculate the MOD 103 required by UCC/EAN
' - Interleave the numbers into printable characters
'ReturnType 0 returns data formatted to the barcode font
     If ReturnType = 0 Then SCC14 = UCC128("01" & DataToEncode & CheckDigit)
'ReturnType 1 returns data formatted for human readable text
     If ReturnType = 1 Then SCC14 = "(01) " & Mid(DataToEncode, 1, 1) & " " & Mid(DataToEncode, 2, 7) & " " & Mid(DataToEncode, 9, 5) & " " & CheckDigit
'ReturnType 2 returns the MOD10 check digit for the data supplied
     If ReturnType = 2 Then SCC14 = STR(CheckDigit)
End Function


Public Function UCC128(DataToEncode As String) As String
'
' This code is Copyright, IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com
' The purpose for this code is to generate a check digit and print a barcode
' according to the UCC-128 EAN-128, SSCC-18 and SCC-14 standards.
'
' UCC/EAN-128 calls for the FNC1 character to be entered, since this cannot
' be printed from the keyboard you must enter FA for the FNC1 code.
' The first FNC1 code is included automatically but you may need to enter this FA
' code if you need to enter another FNC1 code in the middle of the number.
' If you do this MAKE SURE that EVEN numbers are between "FA"; this code performs
' no checking for this!!
'
' Here is an example:  1234FA567800
'
' You MUST use the fully functional Code 128 (dated 12/2000 or later)
' font for this code to create and print a proper barcode
'
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
' Check to make sure data is numeric or "FA" and remove all others.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength Step 2
    'Add all numbers and "FA" to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 2)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 2)
          If Mid(DataToEncode, i, 2) = "FA" Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 2)
     Next i
     DataToEncode = OnlyCorrectData
'Assign start, stop and FNC1 codes
     StartCode = Chr(205)
     StopCode = Chr(206)
     Fnc1 = Chr(202)
' CurrentValue
'<<<< Calculate Modulo 103 Check Digit and generate DataToPrint >>>>
'Set WeightedTotal to the Code 128 value of the start character + Fnc1
     weightedTotal = 105 + 102
     WeightValue = 2
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength Step 2
    'Get the value of each number pair
          CurrentChar = Mid(DataToEncode, i, 2)
    'get the DataToPrint
          If CurrentChar <> "FA" Then
        'set the Integer CurrentValue to the number of String CurrentChar
               CurrentValue = CInt(CurrentChar)
               If CurrentValue < 95 And CurrentValue > 0 Then DataToPrint = DataToPrint & Chr(CurrentValue + 32)
               If CurrentValue > 94 Then DataToPrint = DataToPrint & Chr(CurrentValue + 100)
               If CurrentValue = 0 Then DataToPrint = DataToPrint & Chr(194)
          Else
               If CurrentChar = "FA" Then DataToPrint = DataToPrint & Chr(202)
          End If
    'multiply by the weighting character
          If CurrentChar <> "FA" Then CurrentValue = CurrentValue * WeightValue
          If CurrentChar = "FA" Then CurrentValue = 102 * WeightValue
    'add the values together to get the weighted total
          weightedTotal = weightedTotal + CurrentValue
          WeightValue = WeightValue + 1
     Next i
'divide the WeightedTotal by 103 and get the remainder, this is the CheckDigitValue
     CheckDigitValue = (weightedTotal Mod 103)
'Now that we have the CheckDigitValue, find the corresponding ASCII character from the table
     If CheckDigitValue < 95 And CheckDigitValue > 0 Then C128_CheckDigit = Chr(CheckDigitValue + 32)
     If CheckDigitValue > 94 Then C128_CheckDigit = Chr(CheckDigitValue + 100)
     If CheckDigitValue = 0 Then C128_CheckDigit = Chr(194)
'Get Printable String
     Printable_string = StartCode & Fnc1 & DataToPrint & C128_CheckDigit & StopCode & " "
'Return PrintableString
     UCC128 = Printable_string
End Function


Public Function Code11(DataToEncode As String) As String
'
' Copyright © IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com or http://www.IDautomation.com
'
' You may use our source code in your applications only if you are using barcode fonts created by IDautomation.com, Inc.
' and you do not remove the copyright notices in the source code.
'
' The purpose of this code is to calculate the Code 11 barcode
' Enter all the numbers without dashes
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
' Check to make sure data is numeric or a dash and remove all others.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
          If Mid(DataToEncode, i, 1) = "-" Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
     DataToEncode = OnlyCorrectData
'<<<< Calculate Check Digit >>>>
     Factor = 1
     weightedTotal = 0
     For i = Len(DataToEncode) To 1 Step -1
    'Get the value of each number starting at the end
          CurrentChar = Mid(DataToEncode, i, 1)
    'Set the "-" character to the value of 10
          If CurrentChar = "-" Then CurrentChar = "10"
    'multiply by the weighting character and add together
          weightedTotal = weightedTotal + (Val(CurrentChar) * Factor)
    'change factor for next calculation
          Factor = Factor + 1
     Next i
'Find the Modulo 11 check digit
     CheckDigit = (weightedTotal Mod 11)
'Get Printable String
     Printable_string = "(" & DataToEncode & CheckDigit & ")" & " "
'Return the PrintableString
     Code11 = Printable_string
End Function


Public Function RM4SCC(DataToEncode As String) As String
'
' This module is Copyright, IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com
'
' The purpose of this code is to print the CODE 39 barcode
' You MUST install the AdvRM font for this application to print
'
' Get data from user, this is the DataToEncode
     DataToEncode = RTrim(LTrim(DataToEncode))
     DataToEncode = UCase(DataToEncode)
'only pass correct data
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Get each character one at a time
          CurrentCharNum = (Asc(Mid(DataToEncode, i, 1)))
    'Get the value of CurrentChar according to MOD43
    '0-9
          If CurrentCharNum < 58 And CurrentCharNum > 47 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
    'A-Z
          If CurrentCharNum < 91 And CurrentCharNum > 64 Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
     DataToEncode = OnlyCorrectData
     DataToPrint = DataToEncode
     
     Dim r As Integer
     Dim C As Integer
     Dim Rtotal As Long
     Dim Ctotal As Long
     Rtotal = 0
     Ctotal = 0
     weightedTotal = 0
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Get each character one at a time
          CurrentChar = Mid(DataToEncode, i, 1)
    'Get the values of CurrentChar
          Select Case CurrentChar
          Case "0"
               r = 1
               C = 1
          Case "1"
               r = 1
               C = 2
          Case "2"
               r = 1
               C = 3
          Case "3"
               r = 1
               C = 4
          Case "4"
               r = 1
               C = 5
          Case "5"
               r = 1
               C = 0
          Case "6"
               r = 2
               C = 1
          Case "7"
               r = 2
               C = 2
          Case "8"
               r = 2
               C = 3
          Case "9"
               r = 2
               C = 4
          Case "A"
               r = 2
               C = 5
          Case "B"
               r = 2
               C = 0
          Case "C"
               r = 3
               C = 1
          Case "D"
               r = 3
               C = 2
          Case "E"
               r = 3
               C = 3
          Case "F"
               r = 3
               C = 4
          Case "G"
               r = 3
               C = 5
          Case "H"
               r = 3
               C = 0
          Case "I"
               r = 4
               C = 1
          Case "J"
               r = 4
               C = 2
          Case "K"
               r = 4
               C = 3
          Case "L"
               r = 4
               C = 4
          Case "M"
               r = 4
               C = 5
          Case "N"
               r = 4
               C = 0
          Case "O"
               r = 5
               C = 1
          Case "P"
               r = 5
               C = 2
          Case "Q"
               r = 5
               C = 3
          Case "R"
               r = 5
               C = 4
          Case "S"
               r = 5
               C = 5
          Case "T"
               r = 5
               C = 0
          Case "U"
               r = 0
               C = 1
          Case "V"
               r = 0
               C = 2
          Case "W"
               r = 0
               C = 3
          Case "X"
               r = 0
               C = 4
          Case "Y"
               r = 0
               C = 5
          Case "Z"
               r = 0
               C = 0
               
          End Select
    'add the values together
          Rtotal = Rtotal + r
          Ctotal = Ctotal + C
     Next i
     
'divide the Totals by 6 and get the remainder, this is a reference
'to the Check Digit.
'set check digit to CurrentChar (a string)
     Rtotal = (Rtotal Mod 6)
     Ctotal = (Ctotal Mod 6)
     Select Case Rtotal
     Case 1
          Select Case Ctotal
          Case 1
               CurrentChar = "0"
          Case 2
               CurrentChar = "1"
          Case 3
               CurrentChar = "2"
          Case 4
               CurrentChar = "3"
          Case 5
               CurrentChar = "4"
          Case 0
               CurrentChar = "5"
          End Select
     Case 2
          Select Case Ctotal
          Case 1
               CurrentChar = "6"
          Case 2
               CurrentChar = "7"
          Case 3
               CurrentChar = "8"
          Case 4
               CurrentChar = "9"
          Case 5
               CurrentChar = "A"
          Case 0
               CurrentChar = "B"
          End Select
     Case 3
          Select Case Ctotal
          Case 1
               CurrentChar = "C"
          Case 2
               CurrentChar = "D"
          Case 3
               CurrentChar = "E"
          Case 4
               CurrentChar = "F"
          Case 5
               CurrentChar = "G"
          Case 0
               CurrentChar = "H"
          End Select
     Case 4
          Select Case Ctotal
          Case 1
               CurrentChar = "I"
          Case 2
               CurrentChar = "J"
          Case 3
               CurrentChar = "K"
          Case 4
               CurrentChar = "L"
          Case 5
               CurrentChar = "M"
          Case 0
               CurrentChar = "N"
          End Select
     Case 5
          Select Case Ctotal
          Case 1
               CurrentChar = "O"
          Case 2
               CurrentChar = "P"
          Case 3
               CurrentChar = "Q"
          Case 4
               CurrentChar = "R"
          Case 5
               CurrentChar = "S"
          Case 0
               CurrentChar = "T"
          End Select
     Case 0
          Select Case Ctotal
          Case 1
               CurrentChar = "U"
          Case 2
               CurrentChar = "V"
          Case 3
               CurrentChar = "W"
          Case 4
               CurrentChar = "X"
          Case 5
               CurrentChar = "Y"
          Case 0
               CurrentChar = "Z"
          End Select
     End Select
'Get Printable String
     Printable_string = "(" & DataToPrint & CurrentChar & ")" & " "
'Return PrintableString
     RM4SCC = Printable_string
End Function


Public Function Codabar(DataToEncode As String) As String
'
' This module is Copyright, IDautomation.com, Inc. 2001.  All rights reserved.
' For more info visit http://www.BizFonts.com
'
' The purpose of this code is to print the Codabar barcode
     DataToPrint = ""
     DataToEncode = RTrim(LTrim(DataToEncode))
     
' Check to make sure data is numeric, $, +, -, /, or :, and remove all others.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
    'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
          If Mid(DataToEncode, i, 1) = "$" Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
          If Mid(DataToEncode, i, 1) = "+" Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
          If Mid(DataToEncode, i, 1) = "-" Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
          If Mid(DataToEncode, i, 1) = "/" Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
          If Mid(DataToEncode, i, 1) = "." Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
          If Mid(DataToEncode, i, 1) = ":" Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
     DataToPrint = OnlyCorrectData
'Get Printable String
     Printable_string = "A" & DataToPrint & "B" & " "
'Return PrintableString
     Codabar = Printable_string
End Function


Public Function MOD10(DataToEncode As String) As String
' This is a general MOD10 function like the one required for EAN and UPC
' Check to make sure data is numeric and remove dashes, etc.
     OnlyCorrectData = ""
     StringLength = Len(DataToEncode)
     For i = 1 To StringLength
        'Add all numbers to OnlyCorrectData string
          If IsNumeric(Mid(DataToEncode, i, 1)) Then OnlyCorrectData = OnlyCorrectData & Mid(DataToEncode, i, 1)
     Next i
'<<<< Generate MOD 10 check digit >>>>
     Factor = 3
     weightedTotal = 0
     StringLength = Len(DataToEncode)
     For i = StringLength To 1 Step -1
    'Get the value of each number starting at the end
          CurrentCharNum = Mid(DataToEncode, i, 1)
    'multiply by the weighting factor which is 3,1,3,1...
    'and add the sum together
          weightedTotal = weightedTotal + CurrentCharNum * Factor
    'change factor for next calculation
          Factor = 4 - Factor
     Next i
'Find the CheckDigit by finding the smallest number that = a multiple of 10
     i = (weightedTotal Mod 10)
     If i <> 0 Then
          CheckDigit = (10 - i)
     Else
          CheckDigit = 0
     End If
     MOD10 = STR(CheckDigit)
End Function

