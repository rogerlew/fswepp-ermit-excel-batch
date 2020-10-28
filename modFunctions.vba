Public Function FindMin(arr() As Variant) As Variant

'find minimum value in an array

Dim r As Long
Dim val As Variant
Dim minval As Single
Dim rMin As Long
Dim stat(1) As Variant

minval = 999
For r = 0 To UBound(arr)
     val = arr(r)
     If IsNumeric(val) Then
          If val < minval Then
               minval = val
               rMin = r
          End If
     End If
Next r

stat(0) = rMin
stat(1) = minval

FindMin = stat

End Function
Public Function GetIP() As String

'================================================
'get IP address of computer by calling climate lister perl file
'================================================

Dim ip As String
Dim xml As String
Dim s As String
Dim data() As String
Dim sline As String
Dim p As String
Dim url As String

On Error GoTo EH

url = "https://forest.moscowfsl.wsu.edu/cgi-bin/fswepp/rc/climatefilestsv.pl"

Set xmlhttp = CreateObject("MSXML2.XMLHTTP")

xmlhttp.Open "POST", url, False
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xmlhttp.send (postdata)
xml = xmlhttp.responseText

Set xmlhttp = Nothing

data = Split(xml, Chr(10))
s = data(2)

data = Split(s, " ")
s = data(UBound(data))
ip = Replace(s, "_", ".")

GetIP = ip
Exit Function

EH:
     
     MsgBox ("Could not retrieve computer's IP address")
     GetIP = ""
End Function
Public Function FindMax(arr() As Variant) As Variant

'find maximum value in an array

Dim r As Long
Dim val As Variant
Dim maxval As Single
Dim rMax As Long
Dim stat(1) As Variant

maxval = -999
For r = 0 To UBound(arr)
     val = arr(r)
     If IsNumeric(val) Then
          If val > maxval Then
               maxval = val
               rMax = r
          End If
     End If
Next r

stat(0) = rMax
stat(1) = maxval
FindMax = stat
End Function

Public Function CalcSum(arr() As Variant) As Single

'calculate sum of array values

Dim r As Long
Dim val As Variant
Dim sum As Single

sum = 0
For r = 0 To UBound(arr)
     val = arr(r)
     If IsNumeric(val) Then
          sum = sum + val
     End If
Next r

CalcSum = sum
End Function


Sub SortArray(arr() As Variant)

'================================================================
'sort array in ascending order by bubble sort method
'================================================================

Dim n As Long
Dim temp As Single
Dim nStop As Long
Dim nLast As Long

nStop = UBound(arr) - 1

Do While nStop > i
     nLast = 0
     For n = 0 To nStop
          If arr(n) > arr(n + 1) Then
               temp = arr(n)
               arr(n) = arr(n + 1)
               arr(n + 1) = temp
               nLast = n
          End If
     Next n
     nStop = nLast
Loop

'For n = 0 To UBound(arr)
 '  Debug.Print arr(n)
'Next n

End Sub

Public Function CalcPercentile(data() As Variant, pVal As Integer) As Single

'calculate percentile by interpolation of closest ranks method

Dim n As Integer
Dim p() As Single
Dim n_i As Integer
Dim plow As Single
Dim phigh As Single
Dim vlow As Single
Dim vhigh As Single
Dim v As Single
Dim m As Integer
Dim per As Single

n = UBound(data)
ReDim p(n)

'calculate percentile for each array value
For n_i = 0 To n
      v = data(n_i)
      p(n_i) = (100 / (n + 1)) * (n_i + 1 - 0.5)
      'Debug.Print v, p(n_i)
Next n_i

Select Case pVal
            Case Is <= p(0)
                  per = data(0)
                  
            Case Is >= p(n)
                  per = data(n)
                  
            Case Else
                  For n_i = 0 To n
                        v = data(n_i)
                        
                        Select Case p(n_i)
                              Case Is < pVal
                                   plow = p(n_i)
                                   vlow = v
                                   
                              Case pVal
                                   per = v
                                   Exit For
                              Case Is > pVal
                                   phigh = p(n_i)
                                   vhigh = v
                                   per = vlow + (pVal - plow) / (phigh - plow) * (vhigh - vlow)
                                   Exit For
                         End Select
                    Next n_i
End Select


CalcPercentile = per


End Function

Function FindLastRow(sheetName As String, rStart As Long, col As Integer) As Long

Dim rLast As Long

rLast = rStart
Do Until Sheets(sheetName).Cells(rLast, col) = ""
     rLast = rLast + 1
Loop
FindLastRow = rLast - 1
End Function

Function FindLastColumn(sheetName As String, colStart As Integer, r As Integer) As Integer

Dim colLast As Integer

colLast = colStart
Do Until Sheets(sheetName).Cells(r, colLast) = ""
     colLast = colLast + 1
Loop
FindLastColumn = colLast - 1
End Function

Function ConvertColLetterToNumber(letter As String) As Integer

Dim s As String
Dim colNum As Integer
Dim first As String
Dim second As String

'converts excel column letter  to the corresponding number
'for example column letter "B" = 2

s = Trim(UCase(letter))

Select Case Len(letter)
     Case 1
          colNum = Asc(letter) - 64
          
     Case 2
          first = Left(letter, 1)
          second = Right(letter, 1)
           
          'find number for first letter
          colNum = (Asc(first) - 64) * 26 + Asc(second) - 64
          
End Select

ConvertColLetterToNumber = colNum

End Function
Function ConvertCellAddress(sAdd As String) As Variant
'converts excel cell range to row & column numbers
' for example cell address A1:B4 would be converted to 1,1,4,2

Dim a(3) As Variant
Dim cellsplit() As String

If InStr(1, sAdd, ":") Then
     sAdd = Replace(sAdd, ":", "")
End If

cellsplit = Split(sAdd, "$")
a(0) = cellsplit(2)      'row number
a(1) = ConvertColLetterToNumber(cellsplit(1))      'column number

If UBound(cellsplit) = 4 Then
     a(2) = cellsplit(4)
     a(3) = ConvertColLetterToNumber(cellsplit(3))
Else
     a(2) = 0
     a(3) = 0
End If

ConvertCellAddress = a

End Function

Function ConvertColNumToLetter(colNum As Integer) As String

Dim colLetter As String
Dim x As Long
Dim y As Long

If colNum < 27 Then
     colLetter = Chr(64 + colNum)
Else
     If colNum Mod 26 = 0 Then
          x = colNum / 26 - 1
          y = 26
     Else
          x = Floor(colNum / 26)
          y = colNum - x * 26
     End If
     
     colLetter = Chr(64 + x) & Chr(64 + y)
End If

ConvertColNumToLetter = colLetter
End Function

Function Floor(num As Single) As Long

Dim lngNum As Long

lngNum = CLng(num)

If lngNum > num Then
     lngNum = lngNum - 1
End If

Floor = lngNum
End Function

Function ConvertStringToNumber(s As String) As Double

Dim j As Integer
Dim sNew As String
Dim sChar As String
Dim num As Single

For j = 1 To Len(s)
     sChar = Mid(s, j, 1)
     If IsNumeric(sChar) Or sChar = "." Then
          sNew = sNew & sChar
     End If
Next j

If sNew = "" Then
     num = -999
Else
     num = CSng(sNew)
End If

ConvertStringToNumber = num

End Function

Function GetProbabilityArray(hs_ID As Long, Treatment As Integer, yr As Integer) As Variant

'find probability array values associated with hillslope, treatment type, and year after fire

Dim rFind As Long
Dim arr As Variant

With Sheets("Probability")
     rFind = ProbStartRow
     
     Do Until .Cells(rFind, 1) = ""
          If .Cells(rFind, 1) = hs_ID And .Cells(rFind, 2) = Treatment And .Cells(rFind, 3) = yr Then
               arr = Sheets("Probability").Range("D" & rFind & ":IA" & rFind)
               Exit Do
            End If
            rFind = rFind + 1
     Loop
     
End With

GetProbabilityArray = arr

End Function

Function GetClimates() As Integer

'****************************************************************
'get climates associated with the computer's IP address
'****************************************************************

Dim xml As String
Dim s As String
Dim data() As String
Dim i As Integer
Dim sline As String
Dim clim() As String
Dim l As Long
Dim address As String
Dim sName As String
Dim count As Integer
Dim Climates(500, 1) As Variant
Dim xmlhttp As Object
Dim p As String
Dim myIP As String
Dim url As String
Dim n As Name
Dim bln As Boolean
Dim r1c1 As String

On Error GoTo EH

myIP = GetIP

If myIP = "" Then
     GetClimates = 0
     Exit Function
End If

'clear climates
With Sheets("Ref")
     r = ClimateStartRow
     Do Until .Cells(r, 1) = ""
          .Cells(r, 1) = ""
          .Cells(r, 2) = ""
          r = r + 1
     Loop
End With

p = Sheets("Inputs").cmbPersonality.Value

'set url string &cb is added to the end of the url to override caching
If p = "" Then
     url = "https://forest.moscowfsl.wsu.edu/cgi-bin/fswepp/rc/climatefilestsv.pl?cb=" & Timer() * 100
Else
     url = "https://forest.moscowfsl.wsu.edu/cgi-bin/fswepp/rc/climatefilestsv.pl?ip=" & myIP & "&me=" & p & "&cb=" & Timer() * 100
End If

Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
xmlhttp.Open "GET", url, False
xmlhttp.setRequestHeader "Content-Type", "text/xml"

xmlhttp.send ""
xml = xmlhttp.responseText

Set xmlhttp = Nothing

data = Split(xml, Chr(10))

For i = 0 To UBound(data)
     
     sline = Trim(data(i))
     If InStr(1, sline, "../working") > 0 Then
               
               clim = Split(sline, "*")
               address = clim(0)
               sName = "*" & clim(1)
               l = Len(address)
               address = Mid(address, 1, l - 1)
               Climates(count, 0) = address
               Climates(count, 1) = sName
               count = count + 1
     End If
Next i

If count > 0 Then
     With Sheets("Ref")
          r = ClimateStartRow
          For i = 0 To count - 1
               .Cells(r, 1) = Climates(i, 0)
               .Cells(r, 2) = Climates(i, 1)
               r = r + 1
          Next i

     End With
     
     'set named climate range
     Sheets("Inputs").cmbClimate.ListFillRange = ""
     For Each n In ActiveWorkbook.Names
          If n.Name = "_climates" Then
               n.Delete
               
               'r1c1 = "Ref!R" & ClimateStartRow & "C2:R" & ClimateStartRow + count - 1 & "C2"
               ActiveWorkbook.Names.Add Name:="_climates", RefersTo:="=Ref!$B$" & ClimateStartRow & ":$B$" & ClimateStartRow + count - 1
               Sheets("Inputs").cmbClimate.ListFillRange = "_climates"
               Exit For
          End If
     Next n
End If

GetClimates = count
Exit Function

EH:
     If Err.Number = -2146697211 Then
          MsgBox "A connection could not be made to retrieve climate information, verify that your internet connection is working"
     Else
          MsgBox "The following error occurred while trying to retrieve the custom climate information: " & vbCrLf & Err.Number & ": " & Err.Description
     End If
     GetClimates = 0
End Function

Function FindSed(ID As Long, pFind As Single, Treatment As Integer, yr As Integer) As Single

'find sediment delivery value corresponding to a given probability
'inputs
 ' 1 - desired probabiliity
 '2 - sediment delivery array
 '3- probability array

Dim p1 As Single
Dim p2 As Single
Dim sed_i1 As Integer
Dim sed_i2 As Integer
Dim sed1 As Single
Dim sed2 As Single
Dim p As Single
Dim sed As Single
Dim factor As Single
Dim diff_min As Single
Dim diff As Single
Dim SedArray As Variant
Dim ProbArray As Variant
Dim rsSed As Long
Dim pmax As Single
Dim sedlim As Integer


sed_conv = Sheets("StoredRunValues").Cells(ID + 6, 3)

rSed = ID + SedStartRow - 1
SedArray = Sheets("Sediment").Range("A" & rSed & ":" & "GR" & rSed)     'read sediment delivery data into array
ProbArray = GetProbabilityArray(ID, Treatment, yr)

'find last numeric value in sedarray
For sedlim = 199 To 2 Step -1
      If SedArray(1, sedlim) <> "" Then
            Exit For
      End If
Next sedlim

'find maximum probability
pmax = ProbArray(1, sedlim)

'if probability to find is greater than maximum probability return last sediment value in array
If pFind > pmax Then
      sed = SedArray(1, sedlim)
      GoTo 10
End If

p1 = 0
p2 = 0
                         
For p_i = 1 To 200
          p = ProbArray(1, p_i)
                         
          Select Case p
                    Case pFind
                         sed = SedArray(1, p_i)
                         GoTo 10             'exact match skip interpolation
                                   
                    Case Is < pFind
                         If p > p1 Then
                              sed_i1 = p_i
                              p1 = p
                         End If
                              
                    Case Is > pFind
                         p2 = p
                         sed_i2 = p_i
                         Exit For
          End Select
                              
          
Next p_i
          
diff = Abs(p2 - p1)
'interpolate sediment
If diff < 0.0001 Then
          sed = SedArray(1, sed_i1)
Else
     'calculate interpolation factor
     factor = (p2 - pFind) / diff
                              
     If sed_i2 = 1 Then
               sed = SedArray(1, 1)
     Else
               sed1 = SedArray(1, sed_i1)
               sed2 = SedArray(1, sed_i2)
               sed = sed2 - (sed2 - sed1) * factor
     End If
End If
                         
10:
     sed = sed * sed_conv

     FindSed = sed

End Function


Function CalcSedByTreatment(ID As Long, pFind As Single, Treatment As Integer) As Variant

'calculate sediment delivery for the five years after fire
'inputs - hillslope, probability and treatment type (1 to 6)
'calculate sediment delivery for the 6 standard treatment types
'inputs - hillslope, probability and year after fire

Dim SedResults(4) As Variant
Dim y As Integer

For y = 1 To 5
          SedResults(y - 1) = FindSed(ID, pFind, Treatment, y)
Next y

CalcSedByTreatment = SedResults

End Function
Function CalcSedByYear(ID As Long, pFind As Single, yr As Integer) As Variant

'calculate sediment delivery for the 6 standard treatment types
'inputs - hillslope, probability and year after fire

Dim sedval As Single
Dim SedResults(5) As Variant
Dim t_i As Integer

For t_i = 0 To 5
     SedResults(t_i) = FindSed(ID, pFind, t_i, yr)
Next t_i

CalcSedByYear = SedResults

End Function
Function CalcSedForAllSlopes(pFind As Single, Treatment As Integer, yr As Integer) As Variant

'calculate sediment delivery for all hillslopes for a given probability, treatment and year

Dim hs_count As Long
Dim hs_i As Long
Dim SedResults() As Variant

hs_count = Sheets("Ref").[cpcount]
ReDim SedResults(hs_count - 1)

'loop through hill slopes
For hs_i = 0 To UBound(SedResults)
     SedResults(hs_i) = FindSed(hs_i + 1, pFind, Treatment, yr)
Next hs_i

CalcSedForAllSlopes = SedResults

End Function

Function GetClimateAddress(ClimateName As String) As String

Dim rCell As Range
Dim addr As String
For Each rCell In Sheets("Ref").[_climates]
     If rCell.Value = ClimateName Then
          addr = Sheets("Ref").Cells(rCell.row, rCell.Column - 1)
          Exit For
     End If
Next rCell

GetClimateAddress = addr
End Function


