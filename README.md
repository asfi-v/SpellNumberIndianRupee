# SpellINR VBA Function

## Description
This repository contains a VBA function `SpellINR` that converts a numeric value into its equivalent in Indian Rupees and Paise in words. This function is particularly useful for financial applications where amounts need to be displayed in words.

## Function Overview
This function takes a numeric value as input and returns a string representing the amount in Indian Rupees and Paise in words.

### Features

- Converts numbers into words for Indian Rupees and Paise.
- Handles amounts up to Crores.
- Displays "No Rupees" for zero amounts.
- Customizable for different large number names (e.g., Lakh, Crore).

## Usage
1. Open your Excel workbook.
2. Press `ALT + F11` to open the VBA editor.
3. Click the Insert tab, and click Module.

   ![image](https://github.com/user-attachments/assets/95c403c8-f820-4924-9f91-9f24a02ba287)
   
5. Copy the below VBA Code in Module.
6. Use the `SpellINR` function in your Excel formulas.

   ![image](https://github.com/user-attachments/assets/3c9ec07c-2d6c-4a82-bd6d-499e82282af1)


## VBA Code:
```vba

Function SpellINR(ByVal MyNumber)
    '****************' Main Function *'****************

    Dim Rupees, Paise, Temp
    Dim DecimalPlace, Count
    ReDim Place(9) As String
    Place(2) = " Thousand, "
    Place(3) = " Lakh, "
    Place(4) = " Crore, "
    Place(5) = " Thousand, "
    MyNumber = Trim(Str(MyNumber)) ' Position of decimal place 0 if none
    DecimalPlace = InStr(MyNumber, ".")
    ' Convert Paise and set MyNumber to Rupee amount
    If DecimalPlace > 0 Then
        Paise = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    Count = 1
    Do While MyNumber <> ""
        If Count = 1 Then Temp = GetHundreds(Right(MyNumber, 3))
        If Count > 1 Then Temp = GetHundreds(Right(MyNumber, 2))
        If Temp <> "" Then Rupees = Temp & Place(Count) & Rupees
        If Count = 1 And Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            If Count > 1 And Len(MyNumber) > 2 Then
                MyNumber = Left(MyNumber, Len(MyNumber) - 2)
            Else
                MyNumber = ""
            End If
        End If
        Count = Count + 1
    Loop
    ' Remove the trailing comma and space
    If Right(Rupees, 2) = ", " Then
        Rupees = Left(Rupees, Len(Rupees) - 2)
    End If
    Select Case Rupees
    Case ""
        Rupees = "No Rupees"
    Case "One"
        Rupees = "One Rupee"
    Case Else
        Rupees = "Rupees " & Rupees
    End Select
    Select Case Paise
    Case ""
        Paise = ""
    Case "One"
        Paise = " and One Paisa"
    Case Else
        Paise = " and " & Paise & " Paise"
    End Select
    ' Handle special cases
    If Rupees = "No Rupees" And Paise <> "" Then
        SpellINR = "Rupees " & Mid(Paise, 6) & " Only" ' Remove " and " from Paise
    ElseIf Rupees = "No Rupees" And Paise = "" Then
        SpellINR = "No Rupees Only"
    Else
        SpellINR = Rupees & Paise & " Only"
    End If
End Function

'*******************************************
' Converts a number from 100-999 into text *
'*******************************************
Function GetHundreds(ByVal MyNumber)
    Dim Result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3) 'Convert the hundreds place
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
    End If
    'Convert the tens and ones place
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
    GetHundreds = Result
End Function

'*********************************************
' Converts a number from 10 to 99 into text. *
'*********************************************
Function GetTens(TensText)
    Dim Result As String
    Result = "" ' null out the temporary function value

    If Val(Left(TensText, 1)) = 1 Then ' If value between 10-19
        Select Case Val(TensText)
        Case 10: Result = "Ten"
        Case 11: Result = "Eleven"
        Case 12: Result = "Twelve"
        Case 13: Result = "Thirteen"
        Case 14: Result = "Fourteen"
        Case 15: Result = "Fifteen"
        Case 16: Result = "Sixteen"
        Case 17: Result = "Seventeen"
        Case 18: Result = "Eighteen"
        Case 19: Result = "Nineteen"
        Case Else
        End Select
    Else ' If value between 20-99
        Select Case Val(Left(TensText, 1))
        Case 2: Result = "Twenty "
        Case 3: Result = "Thirty "
        Case 4: Result = "Forty "
        Case 5: Result = "Fifty "
        Case 6: Result = "Sixty "
        Case 7: Result = "Seventy "
        Case 8: Result = "Eighty "
        Case 9: Result = "Ninety "
        Case Else
        End Select
        Result = Result & GetDigit(Right(TensText, 1)) 'Retrieve ones place
    End If
    GetTens = Result
End Function

'*******************************************
' Converts a number from 1 to 9 into text. *
'*******************************************
Function GetDigit(Digit)
    Select Case Val(Digit)
    Case 1: GetDigit = "One"
    Case 2: GetDigit = "Two"
    Case 3: GetDigit = "Three"
    Case 4: GetDigit = "Four"
    Case 5: GetDigit = "Five"
    Case 6: GetDigit = "Six"
    Case 7: GetDigit = "Seven"
    Case 8: GetDigit = "Eight"
    Case 9: GetDigit = "Nine"
    Case Else: GetDigit = ""
    End Select
End Function

```

### Example
1. Type the formula =SpellINR(A1) into the cell where you want to display a written number, where A1 is the cell containing the number you want to convert. You can also manually type the value like =SpellINR(22.500)
2. Press Enter to confirm the formula.
3. Save your SpellINR function workbook.
4. _Excel cannot save a workbook with macro functions in the standard macro-free workbook format (.xlsx). If you click File > Save. A VB project dialog box opens. Click No._

## You can save your file as an **Excel Macro-Enabled Workbook (.xlsm)** to keep your file in its current format.
1. Click **File > Save As**.
2. Click the **Save as type** drop-down menu, and select **Excel Macro-Enabled Workbook**.

   ![image](https://github.com/user-attachments/assets/dcda34a5-7532-4517-8e5b-761cf7997239)
   
4. Click **Save**.

## Click the "Enable Content" buttom. if prompted while opening the "Excel Macro-Enabled Workbook"

   ![image](https://github.com/user-attachments/assets/92e4d3fb-00a7-4d25-b047-6b75fdd627c2)
   
## Acknowledgments
   Inspired by various online resources and examples for converting numbers to words in VBA.
