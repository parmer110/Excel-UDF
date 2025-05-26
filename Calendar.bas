Attribute VB_Name = "Calendar"
Option Explicit

Function CleanTrimText(ByVal txt As String) As String
    ' Combine CLEAN and TRIM functions
    CleanTrimText = Application.WorksheetFunction.Clean(Application.WorksheetFunction.Trim(txt))
End Function ' CleanTrimText

Function CleanAndPrepareDateString(ByVal text As String) As String
    Dim i As Long
    Dim char As String
    Dim cleanedText As String
    Dim preparedText As String

    ' Step 1: Trim leading/trailing spaces
    cleanedText = Trim(text)

    ' Step 2: Remove non-printable characters (similar to WorksheetFunction.Clean)
    For i = 1 To Len(cleanedText)
        char = Mid(cleanedText, i, 1)
        ' Keep printable characters (ASCII >= 32) and Tab (ASCII 9)
        ' You can adjust this condition if you want to keep specific control characters like line feeds (CHAR(10))
        If Asc(char) >= 32 Or Asc(char) = 9 Then
            preparedText = preparedText & char
        End If
    Next i

    ' Step 3: Convert Persian/Arabic digits to Latin digits
    ' Note: If you only use Persian digits, the 0-9 range covers both.
    ' If you might have Arabic (eastern Arabic numerals) too, add them.
    preparedText = Replace(preparedText, "?", "0") ' Persian zero
    preparedText = Replace(preparedText, "?", "1") ' Persian one
    preparedText = Replace(preparedText, "?", "2") ' Persian two
    preparedText = Replace(preparedText, "?", "3") ' Persian three
    preparedText = Replace(preparedText, "?", "4") ' Persian four
    preparedText = Replace(preparedText, "?", "5") ' Persian five
    preparedText = Replace(preparedText, "?", "6") ' Persian six
    preparedText = Replace(preparedText, "?", "7") ' Persian seven
    preparedText = Replace(preparedText, "?", "8") ' Persian eight
    preparedText = Replace(preparedText, "?", "9") ' Persian nine

    ' If you might encounter Eastern Arabic numerals (used in some Arab countries):
    ' preparedText = Replace(preparedText, ChrW(&H6F0), "0") ' Arabic zero
    ' preparedText = Replace(preparedText, ChrW(&H6F1), "1") ' Arabic one
    ' ... and so on for 6F2 to 6F9

    CleanAndPrepareDateString = preparedText
End Function 'CleanAndPrepareDateString

Function NumCvt(ByVal Number As Double) As String
Dim Adad As String

Number = Int(Number)
If Number = 0 Then NumCvt = ChrW(1589) & ChrW(1601) & ChrW(1585): Exit Function

Dim flag As Boolean
Dim S As String
Dim i, L As Byte
Dim k(1 To 5) As Double

S = Trim(str(Number))
L = Len(S)
If L > 15 Then
Adad = ChrW(1576) & ChrW(1587) & ChrW(1610) & ChrW(1575) & ChrW(1585) & ChrW(1576) & ChrW(1586) & ChrW(1585) & ChrW(1711) & " "
Exit Function
End If
For i = 1 To 15 - L
S = "0" & S
Next i
For i = 1 To Int((L / 3) + 0.99)
k(5 - i + 1) = Val(Mid(S, 3 * (5 - i) + 1, 3))
Next i
flag = False
S = " "
For i = 1 To 5
If k(i) <> 0 Then
Select Case i
Case 1
S = S & Three(k(i)) & " " & ChrW(1578) & ChrW(1585) & ChrW(1610) & ChrW(1604) & ChrW(1610) & ChrW(1608) & ChrW(1606) & " "
flag = True
Case 2
S = S & IIf(flag = True, ChrW(1608), " ") & Three(k(i)) & " " & ChrW(1605) & ChrW(1610) & ChrW(1604) & ChrW(1610) & ChrW(1575) & ChrW(1585) & ChrW(1583) & " "
flag = True
Case 3
S = S & IIf(flag = True, ChrW(1608), " ") & Three(k(i)) & " " & ChrW(1605) & ChrW(1610) & ChrW(1604) & ChrW(1610) & ChrW(1608) & ChrW(1606) & " "
flag = True
Case 4
S = S & IIf(flag = True, ChrW(1608), " ") & Three(k(i)) & " " & ChrW(1607) & ChrW(1586) & ChrW(1575) & ChrW(1585) & " "
flag = True
Case 5
S = S & IIf(flag = True, ChrW(1608), " ") & Three(k(i)) & " "
End Select
End If
Next i
NumCvt = Trim(S)
End Function 'NumCvt
Function Three(ByVal Number As Integer) As String
Dim S As String
Dim i, L As Long
Dim H(1 To 3) As Byte
Dim flag As Boolean
L = Len(Trim(str(Number)))
If Number = 0 Then
Three = " "
Exit Function
End If
If Number = 100 Then
Three = ChrW(1610) & ChrW(1603) & ChrW(1589) & ChrW(1583) & " "
Exit Function
End If

If L = 2 Then H(1) = 0
If L = 1 Then
H(1) = 0
H(2) = 0
End If

For i = 1 To L
H(3 - i + 1) = Mid(Trim(str(Number)), L - i + 1, 1)
Next i

Select Case H(1)
Case 1
S = " " & ChrW(1610) & ChrW(1603) & ChrW(1589) & ChrW(1583) & " "
Case 2
S = " " & ChrW(1583) & ChrW(1608) & ChrW(1610) & ChrW(1587) & ChrW(1578) & " "
Case 3
S = " " & ChrW(1587) & ChrW(1610) & ChrW(1589) & ChrW(1583) & " "
Case 4
S = " " & ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585) & ChrW(1589) & ChrW(1583) & " "
Case 5
S = " " & ChrW(1662) & ChrW(1575) & ChrW(1606) & ChrW(1589) & ChrW(1583) & " "
Case 6
S = " " & ChrW(1588) & ChrW(1588) & ChrW(1589) & ChrW(1583) & " "
Case 7
S = " " & ChrW(1607) & ChrW(1601) & ChrW(1578) & ChrW(1589) & ChrW(1583) & " "
Case 8
S = " " & ChrW(1607) & ChrW(1588) & ChrW(1578) & ChrW(1589) & ChrW(1583) & " "
Case 9
S = " " & ChrW(1606) & ChrW(1607) & ChrW(1589) & ChrW(1583) & " "
End Select

Select Case H(2)
Case 1
Select Case H(3)
Case 0
S = S & ChrW(1608) & " " & ChrW(1583) & ChrW(1607) & "  "
Case 1
S = S & ChrW(1608) & " " & ChrW(1610) & ChrW(1575) & ChrW(1586) & ChrW(1583) & ChrW(1607) & " "
Case 2
S = S & ChrW(1608) & " " & ChrW(1583) & ChrW(1608) & ChrW(1575) & ChrW(1586) & ChrW(1583) & ChrW(1607) & " "
Case 3
S = S & ChrW(1608) & " " & ChrW(1587) & ChrW(1610) & ChrW(1586) & ChrW(1583) & ChrW(1607) & " "
Case 4
S = S & ChrW(1608) & " " & ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585) & ChrW(1583) & ChrW(1607) & " "
Case 5
S = S & ChrW(1608) & " " & ChrW(1662) & ChrW(1575) & ChrW(1606) & ChrW(1586) & ChrW(1583) & ChrW(1607) & " "
Case 6
S = S & ChrW(1608) & " " & ChrW(1588) & ChrW(1575) & ChrW(1606) & ChrW(1586) & ChrW(1583) & ChrW(1607) & " "
Case 7
S = S & ChrW(1608) & " " & ChrW(1607) & ChrW(1601) & ChrW(1583) & ChrW(1607) & " "
Case 8
S = S & ChrW(1608) & " " & ChrW(1607) & ChrW(1580) & ChrW(1583) & ChrW(1607) & " "
Case 9
S = S & ChrW(1608) & " " & ChrW(1606) & ChrW(1608) & ChrW(1586) & ChrW(1583) & ChrW(1607) & " "
End Select

Case 2
S = S & ChrW(1608) & " " & ChrW(1576) & ChrW(1610) & ChrW(1587) & ChrW(1578) & " "
Case 3
S = S & ChrW(1608) & " " & ChrW(1587) & ChrW(1610) & " "
Case 4
S = S & ChrW(1608) & " " & ChrW(1670) & ChrW(1607) & ChrW(1604) & " "
Case 5
S = S & ChrW(1608) & " " & ChrW(1662) & ChrW(1606) & ChrW(1580) & ChrW(1575) & ChrW(1607) & " "
Case 6
S = S & ChrW(1608) & " " & ChrW(1588) & ChrW(1589) & ChrW(1578) & " "
Case 7
S = S & ChrW(1608) & " " & ChrW(1607) & ChrW(1601) & ChrW(1578) & ChrW(1575) & ChrW(1583) & " "
Case 8
S = S & ChrW(1608) & " " & ChrW(1607) & ChrW(1588) & ChrW(1578) & ChrW(1575) & ChrW(1583) & " "
Case 9
S = S & ChrW(1608) & " " & ChrW(1606) & ChrW(1608) & ChrW(1583) & " "
End Select

If H(2) <> 1 Then
Select Case H(3)
Case 1
S = S & ChrW(1608) & " " & ChrW(1610) & ChrW(1603)
Case 2
S = S & ChrW(1608) & " " & ChrW(1583) & ChrW(1608)
Case 3
S = S & ChrW(1608) & " " & ChrW(1587) & ChrW(1607)
Case 4
S = S & ChrW(1608) & " " & ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585)
Case 5
S = S & ChrW(1608) & " " & ChrW(1662) & ChrW(1606) & ChrW(1580)
Case 6
S = S & ChrW(1608) & " " & ChrW(1588) & ChrW(1588)
Case 7
S = S & ChrW(1608) & " " & ChrW(1607) & ChrW(1601) & ChrW(1578)
Case 8
S = S & ChrW(1608) & " " & ChrW(1607) & ChrW(1588) & ChrW(1578)
Case 9
S = S & ChrW(1608) & " " & ChrW(1606) & ChrW(1607)
End Select
End If
S = IIf(L < 3, Right(S, Len(S) - 1), S)
Three = S
End Function 'Three

Private Sub CommandButton1_Click()
Dim nowDate As Date
nowDate = Date
 
Dim jalaliDateArray1
jalaliDateArray1 = toJalaaliFromDateObject(nowDate)
 
'MsgBox "Now= " & jalaliDateArray1(0) & "/" & jalaliDateArray1(1) & "/" & jalaliDateArray1(2)

Dim gregorianDate As Date

gregorianDate = toGregorianDateObject(1394, 12, 19)

MsgBox "Date Object From Jalali= " & gregorianDate

Dim jalaliDateArray
jalaliDateArray = toJalaali(2016, 3, 9)

'MsgBox jalaliDateArray(0) & "/" & jalaliDateArray(1) & "/" & jalaliDateArray(2)

Dim gregorialDateArray
gregorialDateArray = toGregorian(1394, 12, 19)

'MsgBox gregorialDateArray(0) & "/" & gregorialDateArray(1) & "/" & gregorialDateArray(2)

End Sub 'CommandButton1_Click

'tabdil as date object be arraye jalali
Private Function toJalaaliFromDateObject(gDate As Date)
   toJalaaliFromDateObject = toJalaali(Year(gDate), Month(gDate), Day(gDate))
End Function 'toJalaaliFromDateObject
' tabdild to object date gregorian
Private Function toGregorianDateObject(jy As Long, jm As Long, jd As Long)
    Dim result
    result = toGregorian(jy, jm, jd)
    toGregorianDateObject = DateValue(result(0) & "-" & result(1) & "-" & result(2))
End Function 'toGregorianDateObject
'tabdil jalali be miladi ba daryaf sale , mahe , rooz jalali
' yek arraye barmigardanad ke index(0)= sal, index(1)=mah, index(2)=rooz
Private Function toGregorian(jy As Long, jm As Long, jd As Long)
      toGregorian = d2g(j2d(jy, jm, jd))
End Function 'toGregorian

' tabdile tarikh miladi be jalali ba daryaft sale, mah , rooz miladi
' yek arraye barmigardanad ke index(0)= sal, index(1)=mah, index(2)=rooz
Private Function toJalaali(gy As Long, gm As Long, gd As Long)
    toJalaali = d2j(g2d(gy, gm, gd))
End Function 'toJalaali
  
' check valid bodan jalali date
Private Function isValidJalaaliDate(jy As Long, jm As Long, jd As Long)
    isValidJalaaliDate = jy >= -61 And jy <= 3177 And jm >= 1 And jm <= 12 And jd >= 1 And jd <= jalaaliMonthLength(jy, jm)
End Function 'isValidJalaaliDate
  

' tedade rooz haye mah ra baraye sale , mahe jalali bar migardanad
Private Function jalaaliMonthLength(jy As Long, jm As Long)
    If (jm <= 6) Then
        jalaaliMonthLength = 31
        Exit Function
    End If
    If (jm <= 11) Then
     jalaaliMonthLength = 30
     Exit Function
    End If
    If (isLeapJalaaliYear(jy)) Then
        jalaaliMonthLength = 30
        Exit Function
        
    End If
  jalaaliMonthLength = 29
End Function 'jalaaliMonthLength
' check in ke sale jalali kabise as ya na
Private Function isLeapJalaaliYear(jy As Long)
    Dim leap As Long
    leap = jalCal(jy)(0)
    
    If (leap = 0) Then
        isLeapJalaaliYear = True
    Else
        isLeapJalaaliYear = False
    End If
    
End Function 'isLeapJalaaliYear

' function haye paeeen baraye amaliate dakheli ast va nabayad estefade shavad
Private Function j2d(jy As Long, jm As Long, jd As Long)
    Dim r As Long
    Dim rgy As Long
    Dim rmarch As Long
    
    rgy = jalCal(jy)(1)
    rmarch = jalCal(jy)(2)
    j2d = g2d(rgy, 3, rmarch) + ((jm - 1) * 31) - ((jm \ 7) * (jm - 7)) + jd - 1
    
End Function 'j2d



Private Function d2j(jdn As Long)
    Dim gy As Long
    gy = d2g(jdn)(0) ' Calculate Gregorian year (gy)
    Dim jy As Long
    jy = gy - 621
    Dim rmarch  As Long
    jalCal (jy)
    rmarch = jalCal(jy)(2)
    Dim rleap  As Long
    rleap = jalCal(jy)(0)
    
    Dim jdn1f As Long
    jdn1f = g2d(gy, 3, rmarch)  'r.march
    Dim jd As Long
    Dim jm As Long
    Dim k As Long

    ' Find number of days that passed since 1 Farvardin.
    k = jdn - jdn1f
    
    Dim result(3)
     
    If (k >= 0) Then
        If (k <= 185) Then
          ' The first 6 months.
          jm = 1 + (k \ 31)
          jd = (k Mod 31) + 1
          result(0) = jy
          result(1) = jm
          result(2) = jd
             
          d2j = result
          Exit Function
          
          
        Else
          ' The remaining months.
          k = k - 186
        End If
    Else
        ' Previous Jalaali year.
        jy = jy - 1
        k = k + 179
        If (rleap = 1) Then 'r.leap
          k = k + 1
        End If
    End If
    
    
    jm = 7 + (k \ 30)
    jd = (k Mod 30) + 1
    
    result(0) = jy
    result(1) = jm
    result(2) = jd
             
    d2j = result
    
End Function 'd2j


Private Function d2g(jdn As Long)
    Dim j As Long
    Dim i As Long
    Dim gd As Long
    Dim gm As Long
    Dim gy As Long
    j = 4 * jdn + 139361631
    j = j + (((((4 * jdn + 183187720) \ 146097) * 3) \ 4) * 4) - 3908
    
    i = (((j Mod 1461) \ 4) * 5) + 308
    gd = ((i Mod 153) \ 5) + 1
    gm = ((i \ 153) Mod 12) + 1
    gy = (j \ 1461) - 100100 + ((8 - gm) \ 6)
    
    Dim result(3)

    result(0) = gy
    result(1) = gm
    result(2) = gd
  
    d2g = result

End Function 'd2g



Private Function g2d(gy As Long, gm As Long, gd As Long)

    Dim d As Long
    d = (((gy + ((gm - 8) \ 6) + 100100) * 1461) \ 4) + ((153 * ((gm + 9) Mod 12) + 2) \ 5) + gd - 34840408
    d = d - ((((gy + 100100 + ((gm - 8) \ 6)) \ 100) * 3) \ 4) + 752
    g2d = d
End Function 'g2d





Private Function jalCal(jy As Long)

    Dim breaks
    breaks = Array(-61, 9, 38, 199, 426, 686, 756, 818, 1111, 1181, 1210, 1635, 2060, 2097, 2192, 2262, 2324, 2394, 2456, 3178)

    Dim bl As Long
    bl = 20
    Dim gy As Long
    
    gy = jy + 621
    Dim leapJ  As Long
    leapJ = -14
    Dim jp As Long
    jp = breaks(0)
    Dim jm As Long
    Dim jump As Long
    Dim leap As Long
    Dim leapG As Long
    Dim march As Long
    Dim n As Long
    Dim i As Long
    

    If (jy < jp Or jy >= breaks(bl - 1)) Then
        MsgBox "Invalid Jalaali year " & jy
    End If
   

   'Find the limiting years for the Jalaali year jy.
   For i = 1 To (bl - 1) Step 1
        jm = breaks(i)
        jump = jm - jp
        If (jy < jm) Then Exit For
        
        leapJ = leapJ + (jump \ 33) * 8 + ((jump Mod 33) \ 4)
        jp = jm
   Next
   
  
   n = jy - jp

  ' Find the number of leap years from AD 621 to the beginning
  ' of the current Jalaali year in the Persian calendar.
  
  leapJ = leapJ + (n \ 33) * 8 + (((n Mod 33) + 3) \ 4)
  If ((jump Mod 33) = 4 And (jump - n) = 4) Then
    leapJ = leapJ + 1
  End If

  ' And the same in the Gregorian calendar (until the year gy).
  leapG = (gy \ 4) - ((((gy \ 100) + 1) * 3) \ 4) - 150

  ' Determine the Gregorian date of Farvardin the 1st.
  march = 20 + leapJ - leapG

  ' Find how many years have passed since the last leap year.
  If ((jump - n) < 6) Then
    n = n - jump + ((jump + 4) \ 33) * 33
  End If
  
  leap = ((((n + 1) Mod 33) - 1) Mod 4)
  If (leap = -1) Then
    leap = 4
  End If
  
  Dim result(3)

  result(0) = leap
  result(1) = gy
  result(2) = march
  
  jalCal = result


End Function 'jalCal


Function S2M(myDate As String)
    
    Dim iYear, iMonth, iDay, iYear1, iMonth1, iDay1 As String
    Dim S, st As String
    Dim i, x As Integer
    Dim arr, arr1 As Variant
    Dim timePartString As String
    Dim spacePos As Long
    Dim dtTime As Date
    Dim iHour As Integer
    Dim iMinute As Integer
    Dim iSecond As Integer
    Dim ExtractTimeFromCombinedString As String
    
    
    ' Clear data
    myDate = CleanAndPrepareDateString(myDate)
          
    For i = 1 To Len(myDate)
        st = Mid(myDate, i, 1)
        
        If (st = "/" Or st = "-" Or st = "." Or st = "\") And x <> 1 Then
            x = 1
            If i = 3 Then S = "13" & S
        Else
            If (st = "/" Or st = "-" Or st = "." Or st = "\") And x = 1 Then
                x = 2
                If i = 5 Or i = 7 Then S = Left(S, Len(S) - 1) & "0" & Right(S, 1)
            Else
                S = S & st
            End If
        End If
    Next i
    If Len(S) = 7 Then S = Left(S, 6) & "0" & Right(S, 1)
    myDate = S
    
    spacePos = InStr(myDate, " ")
    If spacePos > 0 Then
        timePartString = Mid(myDate, spacePos + 1)
        
        dtTime = CDate(timePartString)
        
        iHour = Hour(dtTime)
        iMinute = Minute(dtTime)
        iSecond = Second(dtTime)
        ExtractTimeFromCombinedString = VBA.Format(iHour, "00") & ":" & _
                                        VBA.Format(iMinute, "00") & ":" & _
                                        VBA.Format(iSecond, "00")
    Else
        ExtractTimeFromCombinedString = vbNullString
    End If
    
    
    iYear = Mid(myDate, 1, 4)
    iMonth = Mid(myDate, 5, 2)
    iDay = Mid(myDate, 7, 2)
    If (iDay > 30 And iMonth > 6) Or iDay > 31 Or iMonth > 12 Then S2M = "Error": Exit Function
        
        If (iDay = 30 And iMonth = 12) Then
        iYear1 = iYear + 1
        iMonth1 = 1
        iDay1 = 1
    End If
        
    iYear = Mid(myDate, 1, 4)
    iMonth = Mid(myDate, 5, 2)
    iDay = Mid(myDate, 7, 2)
    If (iDay > 30 And iMonth > 6) Or iDay > 31 Or iMonth > 12 Then S2M = "Error": Exit Function
    
    arr = toGregorian(CInt(iYear), CInt(iMonth), CInt(iDay))
    
    If (iDay = 30 And iMonth = 12) Then
        iYear1 = iYear + 1
        iMonth1 = 1
        iDay1 = 1
        arr1 = toGregorian(CInt(iYear1), CInt(iMonth1), CInt(iDay1))
        If arr1 = arr Then S2M = "Error": Exit Function
    End If
    
    S2M = arr(0) & "/" & Right("0" & arr(1), 2) & "/" & Right("0" & arr(2), 2) & _
        IIf(IsNull(ExtractTimeFromCombinedString), "", " " & ExtractTimeFromCombinedString)
    
   
End Function 'S2M



Function M2S(ByVal myDate As String, Optional Format As Byte = 0)
    
    myDate = CleanAndPrepareDateString(myDate)
   
    ' myDate parameter validation check
    If Not IsDate(myDate) Then
        M2S = CVErr(xlErrValue)
        Exit Function
    End If
    
   
    ' Variables declaration
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    Dim iWeekday As Integer
    
    Dim hasTime As Boolean
    'Dim cellNumberFormat As String
    'cellNumberFormat = myCell.NumberFormatLocal
    
    Dim iHour As Integer
    Dim iMinute As Integer
    Dim iSecond As Integer
    
    Dim mah As String
    Dim rooz As String
    
    Dim persianYear As String
    Dim persianMonth As String
    Dim persianDay As String
    
    Dim dt As Date
    dt = CDate(myDate) ' Sure myDate assuming date/time correctly
    Dim arr As Variant
    
    
    ' Gregorian date part extraction
    iDay = Day(myDate)
    iMonth = Month(myDate)
    iYear = Year(myDate)
    iWeekday = Weekday(myDate)
    
    ' Time part extraction
    iHour = Hour(dt)
    iMinute = Minute(dt)
    iSecond = Second(dt)
    
    If Len(myDate) > 10 Then
        hasTime = True
    Else
        hasTime = False
    End If
    
    ' Convert to Jalali date
    arr = toJalaaliFromDateObject(dt)
    
    iDay = arr(2)
    iMonth = arr(1)
    iYear = str(arr(0))


    Select Case iWeekday
    Case 1
        rooz = ChrW(1740) & ChrW(1705) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607)
    Case 2
        rooz = ChrW(1583) & ChrW(1608) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607)
    Case 3
        rooz = ChrW(1587) & ChrW(1607) & ChrW(8204) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607)
    Case 4
        rooz = ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607)
    Case 5
        rooz = ChrW(1662) & ChrW(1606) & ChrW(1580) & ChrW(8204) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607)
    Case 6
        rooz = ChrW(1580) & ChrW(1605) & ChrW(1593) & ChrW(1607)
    Case 7
        rooz = ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607)
    End Select
    
   
   Select Case iMonth
    
    Case 1
        mah = ChrW(1601) & ChrW(1585) & ChrW(1608) & ChrW(1585) & ChrW(1583) & ChrW(1740) & ChrW(1606)
    Case 2
        mah = ChrW(1575) & ChrW(1585) & ChrW(1583) & ChrW(1740) & ChrW(1576) & ChrW(1607) & ChrW(1588) & ChrW(1578)
    Case 3
        mah = ChrW(1582) & ChrW(1585) & ChrW(1583) & ChrW(1575) & ChrW(1583)
    Case 4
        mah = ChrW(1578) & ChrW(1740) & ChrW(1585)
    Case 5
        mah = ChrW(1605) & ChrW(1585) & ChrW(1583) & ChrW(1575) & ChrW(1583)
    Case 6
        mah = ChrW(1588) & ChrW(1607) & ChrW(1585) & ChrW(1740) & ChrW(1608) & ChrW(1585)
    Case 7
        mah = ChrW(1605) & ChrW(1607) & ChrW(1585)
    Case 8
        mah = ChrW(1570) & ChrW(1576) & ChrW(1575) & ChrW(1606)
    Case 9
        mah = ChrW(1570) & ChrW(1584) & ChrW(1585)
    Case 10
        mah = ChrW(1583) & ChrW(1740)
    Case 11
        mah = ChrW(1576) & ChrW(1607) & ChrW(1605) & ChrW(1606)
    Case 12
        mah = ChrW(1575) & ChrW(1587) & ChrW(1601) & ChrW(1606) & ChrW(1583)
       
    End Select
    
    
    Select Case Format
    
    Case 1
        M2S = rooz & "," & iYear & "/" & VBA.Format(iMonth, "00") & "/" & VBA.Format(iDay, "00")
    Case 2
        M2S = rooz & " " & NumCvt(iDay) & " " & mah & ChrW(8204) & ChrW(1605) & ChrW(1575) & ChrW(1607) & " " & NumCvt(iYear)
    Case 3
        M2S = iYear
    Case 4
        M2S = VBA.Format(iMonth, "00")
    Case 5
        M2S = VBA.Format(iDay, "00")
    Case 6
        M2S = mah
    Case 7
        M2S = rooz
    Case 0
        M2S = iYear & "/" & VBA.Format(iMonth, "00") & "/" & VBA.Format(iDay, "00") & _
            IIf(hasTime, " " & VBA.Format(iHour, "00") & ":" & VBA.Format(iMinute, "00") & ":" & VBA.Format(iSecond, "00"), "")
    Case Else
        M2S = "Only one of this number as format(0,1,2,3,4,5,6,7)"
    End Select
End Function 'M2S

