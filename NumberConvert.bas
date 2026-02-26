Attribute VB_Name = "NumberConvert"
Option Explicit

Private Function GetHundreds(ByVal MyNumber)
    Dim Result As String
    
    If Val(MyNumber) = 0 Then Exit Function
    
    MyNumber = Right("000" & MyNumber, 3)
    
    ' Hundreds place
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
    End If
    
    ' Tens & Ones place
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
    
    GetHundreds = Result
End Function

Private Function GetTens(TensText)
    Dim Result As String
    
    Result = ""
    
    If Val(Left(TensText, 1)) = 1 Then
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
        End Select
    Else
        Select Case Val(Left(TensText, 1))
            Case 2: Result = "Twenty "
            Case 3: Result = "Thirty "
            Case 4: Result = "Forty "
            Case 5: Result = "Fifty "
            Case 6: Result = "Sixty "
            Case 7: Result = "Seventy "
            Case 8: Result = "Eighty "
            Case 9: Result = "Ninety "
        End Select
        
        Result = Result & GetDigit(Right(TensText, 1))
    End If
    
    GetTens = Result
End Function

Private Function GetDigit(Digit)
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
Public Function NumberToWordsEnglish(ByVal MyNumber)
    Dim Dollars, Cents, Temp
    Dim DecimalPlace, Count
    Dim Result As String
    
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    
    MyNumber = Trim(Str(MyNumber))
    DecimalPlace = InStr(MyNumber, ".")
    
    If DecimalPlace > 0 Then
        Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    
    Count = 1
    Do While MyNumber <> ""
        Temp = GetHundreds(Right(MyNumber, 3))
        If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    
    If Dollars = "" Then Dollars = "zero"
    
    Result = LCase(Dollars & " " & Cents)
    
    ' Capitalize first letter only
    NumberToWordsEnglish = UCase(Left(Result, 1)) & Mid(Result, 2)
End Function

Public Function NumberToWordVietnamese(ByVal chuyenso) As String

    Dim s09
    s09 = Array("", " M" & ChrW(7897) & "t", " hai", " ba", " b" & ChrW(7889) & "n", " n" & ChrW(259) & "m", " s" & ChrW(225) & "u", " b" & ChrW(7843) & "y", " t" & ChrW(225) & "m", " ch" & ChrW(237) & "n")
    Dim lop3
    lop3 = Array("", " tri" & ChrW(7879) & "u", " ngh" & ChrW(236) & "n", " t" & ChrW(7927))
    
    Dim dau As String, VND As String
    Dim vt As Long, sonhan As Long
    Dim sochuso As Long
    Dim i As Long, lop As Long
    Dim n1 As Integer, n2 As Integer, n3 As Integer
    Dim baso As String
    Dim s1 As String, s2 As String, s3 As String, s123 As String
    
    If Trim(chuyenso) = "" Then
      VND = ""
    ElseIf IsNumeric(chuyenso) = True Then
    
      If chuyenso < 0 Then dau = ChrW(226) & "m " Else dau = ""
      
      chuyenso = Application.WorksheetFunction.Round(Abs(chuyenso), 0)
      chuyenso = Replace(chuyenso, ",", "")
      
      vt = InStr(1, chuyenso, "E")
      If vt > 0 Then
        sonhan = Val(Mid(chuyenso, vt + 1))
        chuyenso = Trim(Mid(chuyenso, 1, vt - 1))
        chuyenso = chuyenso & String(sonhan - Len(chuyenso) + 1, "0")
      End If
      
      chuyenso = Trim(chuyenso)
      sochuso = Len(chuyenso) Mod 9
      If sochuso > 0 Then chuyenso = String(9 - sochuso, "0") & chuyenso
      
      VND = ""
      i = 1
      lop = 1
      
      Do
        n1 = CInt(Mid(chuyenso, i, 1))
        n2 = CInt(Mid(chuyenso, i + 1, 1))
        n3 = CInt(Mid(chuyenso, i + 2, 1))
        
        i = i + 3
        
        If n1 & n2 & n3 = "000" Then
            s123 = ""
        Else
            
          If n1 = 0 Then
            s1 = ""
          Else
            s1 = s09(n1) & " tr" & ChrW(259) & "m"
          End If
          
          If n2 = 0 Then
            If s1 = "" Or n3 = 0 Then
              s2 = ""
            Else
              s2 = " linh"
            End If
          Else
            If n2 = 1 Then
              s2 = " M" & ChrW(432) & ChrW(7901) & "i"
            Else
              s2 = s09(n2) & " m" & ChrW(432) & ChrW(417) & "i"
            End If
          End If
          
          If n3 = 1 Then
            If n2 <= 1 Then
                s3 = " M" & ChrW(7897) & "t"
            Else
                s3 = " m" & ChrW(7889) & "t"
            End If
          ElseIf n3 = 5 And n2 <> 0 Then
            s3 = " l" & ChrW(259) & "m"
          Else
            s3 = s09(n3)
          End If
          
          If i > Len(chuyenso) Then
            s123 = s1 & s2 & s3
          Else
            s123 = s1 & s2 & s3 & lop3(lop)
          End If
          
        End If
        
        lop = lop + 1
        If lop > 3 Then lop = 1
        
        VND = VND & s123
        
        If i > Len(chuyenso) Then Exit Do
        
      Loop
      
      If VND = "" Then
        VND = "kh" & ChrW(244) & "ng"
      Else
        VND = dau & Trim(VND)
      End If
    
    Else
      VND = chuyenso
    End If
    
    
    VND = LCase(VND)
    
    ' Capitalize first letter only
    NumberToWordVietnamese = UCase(Left(VND, 1)) & Mid(VND, 2)

End Function
