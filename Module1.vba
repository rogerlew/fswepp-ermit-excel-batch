Sub CallErmit(postdata As String)

'****************************************************************
'send input string to ERMIT Perl script and return results as text
'****************************************************************

Dim results As String
Dim xmlhttp As Object

SubName = "CallErmit"

Set xmlhttp = CreateObject("MSXML2.XMLHTTP")

xmlhttp.Open "POST", "https://forest.moscowfsl.wsu.edu/cgi-bin/fswepp/ermit/erm.pl", False
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xmlhttp.send (postdata)
results = xmlhttp.responseText

Set xmlhttp = Nothing

'Open "c:\temp\results.html" For Output As #1
'Print #1, results
'Close #1
ErmitResults = results

End Sub

Sub CheckInputs()

'****************************************************************
'check inputs for errors
'****************************************************************

Dim r As Long
Dim c As Long
Dim count As Integer
Dim LastRow As Long
Dim val As String
Dim ErrMsg As String
Dim ID As Long
Dim maxlen As Integer
Dim grass As String
Dim shrub As String
Dim bare As String
Dim sum As Single
Dim eCount As Long
Dim resp As Integer
Dim E As ErmitInput
Dim sName As String
Dim blnClimateCheck As Boolean

InputErrorCount = 0

Application.ScreenUpdating = False

count = Sheets("Ref").[InputCount]
LastRow = HillslopeStartRow + count - 1

sName = "Inputs"

With Sheets("Inputs")
     .Activate
     
     If .cmbClimateOption.Value = "Multiple Climates" Then
          blnClimateCheck = True
     End If
     
     'clear any cell coloring and comments from previous input check
     Range("B" & HillslopeStartRow & ":Q" & LastRow).Select
     Selection.Interior.ColorIndex = 0
     Selection.ClearComments
     
     Select Case .cboUnits
          Case "English"
               maxlen = 1000
          Case "Metric"
               maxlen = 300
     End Select
     
     For r = HillslopeStartRow To LastRow
            E.hs_code = Trim(.Cells(r, 1))
            E.Area = .Cells(r, 4)
            E.SoilType = Trim(LCase(.Cells(r, 6)))
            E.RockPercent = .Cells(r, 7)
            E.VegType = Trim(.Cells(r, 8))
            E.TopGradient = .Cells(r, 9)
            E.MidGradient = .Cells(r, 10)
            E.ToeGradient = .Cells(r, 11)
            E.ShrubPercent = .Cells(r, 12)
            E.GrassPercent = .Cells(r, 13)
            E.BarePercent = .Cells(r, 14)
            E.HorizontalSlopeLength = .Cells(r, 15)
            E.BurnClass = Trim(LCase(.Cells(r, 17)))
            
            If blnClimateCheck And .Cells(r, 2) = "" Then
               Call AddComment(sName, r, 2, "climate not selected")
            End If
      
          'hillslope code
            If E.hs_code = "" Then
                  Call AddComment(sName, r, 1, "Hillslope code is blank")
          End If
          
          'area
          If IsNumeric(E.Area) Then
               If E.Area <= 0 Then
                    Call AddComment(sName, r, 4, "area must be greater than 0")
               End If
          Else
               Call AddComment(sName, r, 4, "area is blank")
          End If
          
          'soil type
          If E.SoilType = "" Then
                  Call AddComment(sName, r, 6, "soil type is blank")
            Else
                  Select Case E.SoilType
                        Case "silt loam", "loam", "sandy loam", "clay loam"
                        
                        Case Else
                              Call AddComment(sName, r, 6, "soil type must be: silt loam, sandy loam, clay loam, or loam")
                  End Select
            End If

            'rock %
            If IsNumeric(E.RockPercent) Then
                  If E.RockPercent < 0 Or E.RockPercent > 50 Then
                              Call AddComment(sName, r, 7, "Rock % must be between 0 and 50%")
                  End If
            Else
                  Call AddComment(sName, r, 7, "Invalid rock % value")
            End If
                     
             'veg type
            If E.VegType = "" Then
                  Call AddComment(sName, r, 8, "Vegetation type is blank")
            Else
                  Select Case E.VegType
                        Case "Forest", "Range", "Chaparral"
                        
                        Case Else
                               Call AddComment(sName, r, 8, ErrMsg = "Vegetation type must be either: Forest, Range, or Chaparral")
                  End Select
            End If
               
            'gradients
                  If IsNumeric(E.TopGradient) Then
                        If E.TopGradient > 100 Or E.TopGradient < 0 Then
                              Call AddComment(sName, r, 9, "gradient % must be be from 0-100")
                        End If
                  Else
                        Call AddComment(sName, r, 9, "Invalid gradient % value")
                  End If
           
           
           'mid gradient
            If IsNumeric(E.MidGradient) Then
                  If E.MidGradient > 100 Or E.MidGradient < 0 Then
                        Call AddComment(sName, r, 10, "gradient % must be be from 0-100")
                  End If
            Else
                  Call AddComment(sName, r, 10, "Invalid gradient % value")
            End If
            
            'toe gradient
            If IsNumeric(E.ToeGradient) Then
                  If E.ToeGradient > 100 Or E.ToeGradient < 0 Then
                        Call AddComment(sName, r, 11, "gradient % must be be from 0-100")
                  End If
            Else
                  Call AddComment(sName, r, 11, "Invalid gradient % value")
            End If
            
            '% shrub
            If IsNumeric(E.ShrubPercent) Then
                        If E.ShrubPercent < 0 Or E.ShrubPercent > 100 Then
                              Call AddComment(sName, r, 12, "% shrub must be from 0-100")
                        End If
            End If
                    
            '% grass
          If IsNumeric(E.GrassPercent) Then
                        If E.GrassPercent < 0 Or E.GrassPercent > 100 Then
                              Call AddComment(sName, r, 12, "% grass must be from 0-100")
                        End If
            End If
                    
             '% bare
          If IsNumeric(E.BarePercent) Then
                        If E.BarePercent < 0 Or E.BarePercent > 100 Then
                              Call AddComment(sName, r, 13, "% bare must be from 0-100")
                        End If
            End If
            
            'horizontal length
            If IsNumeric(E.HorizontalSlopeLength) Then
                        If E.HorizontalSlopeLength < 1 Or E.HorizontalSlopeLength > maxlen Then
                              Call AddComment(sName, r, 15, "length must be between 1 and " & maxlen & " " & Sheets("Ref").[lengthunits])
                        End If
                          
            Else
                        Call AddComment(sName, r, 15, "Invalid slope length")
            End If
        
            'burn class
            If E.BurnClass = "" Then
                        Call AddComment(sName, r, 17, "burn severity class is blank")
            Else
                  Select Case E.BurnClass
                        Case "low", "moderate", "high", "unburned"      ' DEH 2014.04.08
                        Case Else
                              Call AddComment(sName, r, 17, "burn severity class must be either: High, Medium, Low, or Unburned.")  ' DEH 2014.04.08
                  End Select
            End If

          'check total pre fire %
            If IsNumeric(E.BarePercent) Then
                  sum = E.BarePercent
            End If
            
            If IsNumeric(E.GrassPercent) Then
                  sum = sum + E.GrassPercent
            End If
            
            If IsNumeric(E.ShrubPercent) Then
                  sum = sum + E.ShrubPercent
            End If
          
            If sum > 100 Then
                    ErrMsg = "Pre-fire community total percentage is greater than 100"
                    Call AddComment(sName, r, 12, ErrMsg)
                    Call AddComment(sName, r, 13, ErrMsg)
                     Call AddComment(sName, r, 14, ErrMsg)
            End If
            
            'if vegetation type is forest clear pre-fire % cells
            If E.VegType = "Forest" Then
                        Cells(r, 12) = ""
                        Cells(r, 13) = ""
                        .Cells(r, 14) = ""
            End If
     Next r
     .Cells(HillslopeStartRow, 1).Select
End With

Application.ScreenUpdating = True

If InputErrorCount > 0 Then
     resp = MsgBox(InputErrorCount & " input errors were found.  The cells with errors are highlighted in green." & vbCrLf & "Move the mouse cursor over the highlighted cells to view the error description.", vbOKOnly, "Input Errors")
End If
End Sub

Sub AddComment(sheetName As String, sheetrow As Long, sheetcol As Integer, s As String)

'****************************************************************
'add comment to cell and highlight in green
'****************************************************************
Dim rng As Excel.Range
Dim sAddress As String

sAddress = ConvertColNumToLetter(sheetcol) & sheetrow
InputErrorCount = InputErrorCount + 1

Set rng = Sheets(sheetName).Range(sAddress)


With rng
    .Interior.ColorIndex = 4
     If .Comment Is Nothing Then
     Else
          .Comment.Delete
     End If
     .AddComment Text:=s
End With

Set rng = Nothing

End Sub

Sub RunErmitBatch()

'*******************************************************
'run ermit on a set of hillslopes
'*******************************************************

Dim col As Integer
Dim LastHillslopeRow As Long
Dim InputVariables(14, 1) As String
Dim InputStrings() As String
Dim E As ErmitInput
Dim count As Long
Dim units As String
Dim i As Integer
Dim climate As String
Dim f As String
Dim TimeLeft As String
Dim t1 As Date
Dim t2 As Date
Dim tDiff As Single
Dim RunCount As Integer
Dim HillslopeCount As Long
Dim hs_code As String
Dim total_area As Single
Dim hs_i As Long
Dim RunStatus As Integer
Dim ErrMsg As String
Dim ClimateOption As String
Dim rInput As Integer

'On Error GoTo EH

SetUnits
ClearSheets

frmRun.lblRun.Caption = "Running batch, please wait"
DoEvents

ClimateOption = Sheets("Inputs").cmbClimateOption.Value

HillslopeCount = Sheets("Ref").[InputCount]

' *********** DEH
'call subroutine to log ermit batch
'Call LogErmitRun(HillslopeCount, ClimateOption)
'  --> moved down to capture climate name properly.... 2012.08.16

'loop through hillslopes and create input strings
LastHillslopeRow = HillslopeStartRow + HillslopeCount - 1

RunStatus = 1
With Sheets("Inputs")
     .Activate
     For rInput = HillslopeStartRow To LastHillslopeRow
          hs_code = Trim(.Cells(rInput, 1))
          
          RunCount = RunCount + 1
          ReDim Preserve InputStrings(1 To RunCount)
          total_area = total_area + .Cells(rInput, 4)
          
          Select Case .Cells(rInput, 6)
               Case "Clay Loam"
                    E.SoilType = "clay"
                    
               Case "Silt Loam"
                    E.SoilType = "silt"
                    
               Case "Sandy Loam"
                    E.SoilType = "sand"
                    
               Case "Loam"
                    E.SoilType = "loam"
          End Select
          
         E.RockPercent = .Cells(rInput, 7)
         
         Select Case .Cells(rInput, 8)
               Case "Forest", "Range"
                        E.VegType = LCase(.Cells(rInput, 8))
                        
               Case "Chaparral"
                         E.VegType = "chap"
          End Select
          
          E.TopGradient = .Cells(rInput, 9)
          E.MidGradient = .Cells(rInput, 10)
          E.ToeGradient = .Cells(rInput, 11)
          E.ShrubPercent = .Cells(rInput, 12)
          E.GrassPercent = .Cells(rInput, 13)
          E.BarePercent = .Cells(rInput, 14)
          E.HorizontalSlopeLength = .Cells(rInput, 15)
          E.BurnClass = LCase(Left(.Cells(rInput, 17), 1))
              
          If ClimateOption = "Single Climate" Then
               climate = GetClimateAddress(Sheets("Inputs").cmbClimate.Value)
               ClimateNameD = Sheets("Inputs").cmbClimate.Value
          Else
               climate = GetClimateAddress(.Cells(rInput, 2))
          End If
          
          InputStrings(RunCount) = "?me=&units=" & LenUnits & "&Climate=" & climate & "&SoilType=" & E.SoilType & "&rfg=" & E.RockPercent & "&achtung=Run WEPP&vegetation=" & E.VegType _
                                   & "&top_slope=" & E.TopGradient & "&avg_slope=" & E.MidGradient & "&toe_slope=" & E.ToeGradient & "&length=" & E.HorizontalSlopeLength _
                                   & "&severity=" & E.BurnClass & "&pct_shrub=" & E.ShrubPercent & "&pct_grass=" & E.GrassPercent & "&pct_bare=" & E.BarePercent
                                             
     Next rInput
End With

' *********** DEH *********** moved down 2012.08.16
'call subroutine to log ermit batch run
          If ClimateOption = "Single Climate" Then
               Call LogErmitRun(HillslopeCount, ClimateNameD)
          Else
               Call LogErmitRun(HillslopeCount, ClimateOption)
          End If
' *********** DEH *********** moved down 2012.08.16


Sheets("Summary Results - Year 1").[TotalArea] = CStr(Round(total_area, 2))

RunStatus = 2
'loop through input strings and send to ERMIT
For i = 1 To UBound(InputStrings)
     If blnStopRun = True Then
          Exit Sub
     End If
     
     t1 = Now
     CallErmit (InputStrings(i))
     DoEvents
     GetArrays (i)

     DoEvents
     GetRainfallStats (i)
     GetRunValues (i)
     t2 = Now
     
     tDiff = (t2 - t1) * 86400 * (RunCount - i)
     
     If tDiff < 60 Then
          TimeLeft = CInt(tDiff) & " seconds"
     Else
          TimeLeft = Round(tDiff / 60, 1) & " minutes"
     End If
     
     frmRun.lblRun.Caption = "Running hillslope " & i + 1 & " of " & RunCount & vbCrLf & "Estimated Time Left: " & TimeLeft
     DoEvents
Next i

'calculate year 1 sediment delivery stats for 10, 30, 50 and 75% probabilities for each hillslope
RunStatus = 3
CalcYearOneResults

Exit Sub

EH:
     Select Case RunStatus
          Case 1
               ErrMsg = "The following error occurred while reading the hillslope inputs: " & vbCrLf & Err.Number & "-" & Err.Description
          Case 2
               ErrMsg = "The following error occurred while running hillslope " & i & vbCrLf & Err.Number & "-" & Err.Description
          Case 3
               ErrMsg = "The following error occurred while calculating year 1 results: " & vbCrLf & Err.Number & "-" & Err.Description
          End Select
          
          MsgBox ErrMsg
      
End Sub
Function CalcLog(ID As Long, diam As Single, spacing As Single, yr As Integer, sed0() As Variant) As Variant
' orrected soil type and mid slope reference cells
'   SoilType = LCase(.Cells(InputRow, 5)); slope_mid = .Cells(InputRow, 9)
'*******************************************************
'calculate sediment delivery
'inputs
' hillslope row
' log diameter
' log spacing
' year after fire
' untreated sediment delivery array
'*******************************************************

Dim slope_max  As Single
Dim slope_min As Single
Dim spacing_def As Single
Dim d_def  As Single
Dim slope_mid As Single
Dim sed_bulk_density As Single
Dim eff(4) As Single
Dim caught(4) As Single
Dim logresult(4) As Variant
Dim SoilType As String
Dim row As Integer
Dim d_cm As Single
Dim units As String
Dim capacity_vol As Single
Dim capacity As Single
Dim InputRow As Integer
Dim i10 As Single
Dim i As Integer

'get hillslope inputs
InputRow = HillslopeStartRow + ID - 1
With Sheets("Inputs")
     SoilType = LCase(.Cells(InputRow, 6))
     units = .cboUnits
     slope_mid = .Cells(InputRow, 10)
End With

'peak 10 minute rainfall intensity mm/hr
i10 = Sheets("StoredRunValues").Cells(ID + 6, 2)

slope_max = 100          '%
slope_min = 0.05         '%

If slope_mid > 100 Then
     slope_mid = 100
ElseIf slope_mid < 0.05 Then
     slope_mid = 0.05
End If
     
Select Case units
     Case "English"
          d_def = 1        'ft
          
     Case "Metric"
          d_def = 0.3
End Select
    
'convert diameter to centimeters and spacing to meters
If units = "Metric" Then
     d_cm = diam * 100
Else
     d_cm = diam * 30.48
     spacing = spacing * 0.3048
End If
     
'determine bulk density in g/m^3 based upon soil type
Select Case SoilType
     Case "clay loam"
          sed_bulk_density = 1.1
          
     Case "silt loam"
          sed_bulk_density = 0.97
          
     Case "sandy loam"
          sed_bulk_density = 1.23
          
     Case "loam"
          sed_bulk_density = 1.16
          
     Case Else
          sed_bulk_density = 1
End Select

capacity_vol = coeff_slope / slope_mid + coeff_diam * d_cm ^ 2 + coeff_spacing / spacing + Intercept

If capacity_vol < 0 Then
     capacity_vol = 0
End If

capacity = capacity_vol * sed_bulk_density        'units Mg/hectare

'convert capacity to tons/acre if units are English
If units = "English" Then
     capacity = capacity * 0.4461
End If

'calculate efficiency
eff(0) = 113.97 - 0.8425 * i10
If eff(0) < 0 Then
     eff(0) = 0
ElseIf eff(0) > 100 Then
     eff(0) = 100
End If

eff(1) = 116 - 1.4 * i10
If eff(1) < 0 Then
     eff(1) = 0
ElseIf eff(1) > 100 Then
     eff(1) = 100
End If
    
eff(2) = eff(1) * 0.75
eff(3) = eff(2) * 0.55
eff(4) = eff(3) * 0.45

'calculate caught and result array
For i = 0 To 4
     caught(i) = capacity * eff(i) / 100
     If caught(i) > sed0(i) Then
          caught(i) = sed0(i)
     End If
     logresult(i) = Round(sed0(i) - caught(i), 2)
Next i

CalcLog = logresult
End Function

Sub CalcYearOneResults()

'**************************************************************************************
'calculate summary sediment delivery statistics for year 1 for the six treatment types
' for 10, 30, 50 and 75% probability
'**************************************************************************************
'  don't report percentiles -- too much info!    DEH 2011.08.18
'***********************************
Dim prob(3) As Single
Dim cnt As Integer
Dim i As Integer       'probability index
Dim j As Long       'hillslope index
Dim k As Integer       'treatment index
Dim rStart As Integer
Dim colStart As Integer
Dim row As Integer
Dim SedResults() As Variant
Dim s_i As Integer
Dim hs_min As String
Dim hs_max As String
Dim sed_min As Single
Dim sed_max As Single
Dim sed_avg As Single
Dim sed_25 As Single
Dim sed_50 As Single
Dim sed_75 As Single
Dim sed_sum As Single
Dim sed_hs As Single
Dim r_hs As Long
Dim rc() As Variant

'clear results
With Sheets("Summary Results - Year 1")
     .[_table1] = ""
     .[_table2] = ""
     .[_table3] = ""
     .[_table4] = ""
End With

cnt = Sheets("Ref").[InputCount]
prob(0) = 0.1
prob(1) = 0.3
prob(2) = 0.5
prob(3) = 0.75

ReDim SedResults(cnt - 1)

'loop through probabilities
For i = 0 To 3
     frmRun.lblRun.Caption = "Calculating Year 1 Summary Results" & vbCrLf & "Probability - " & CInt(prob(i) * 100) & " %"
     DoEvents
          
     Select Case i
          Case 0
               rc = ConvertCellAddress(Sheets("Summary Results - Year 1").[_table1].address)
          Case 1
               rc = ConvertCellAddress(Sheets("Summary Results - Year 1").[_table2].address)
          Case 2
               rc = ConvertCellAddress(Sheets("Summary Results - Year 1").[_table3].address)
          Case 3
               rc = ConvertCellAddress(Sheets("Summary Results - Year 1").[_table4].address)
     End Select
          
     rStart = rc(0)
     colStart = rc(1)
     
     'loop through treatment types
     For k = 0 To 5
          SedResults = CalcSedForAllSlopes(prob(i), k, 1)        'call function to calculate sediment delivery for all hillslopes
          
          'calculate min max and average
          sed_min = 999
          sed_max = 0
          sed_sum = 0
          
          For s_i = 0 To UBound(SedResults)
               sed_hs = SedResults(s_i)
               sed_sum = sed_sum + sed_hs
               
               If sed_hs < sed_min Then
                    sed_min = sed_hs
                    hs_min = Sheets("Inputs").Cells(HillslopeStartRow + s_i, 1)
               End If
               
               If sed_hs > sed_max Then
                    sed_max = sed_hs
                    hs_max = Sheets("Inputs").Cells(HillslopeStartRow + s_i, 1)
               End If
          Next s_i
                    
          sed_avg = sed_sum / cnt
          
          'calculate percentiles
'          Call SortArray(SedResults)
'          sed25 = CalcPercentile(SedResults, 25)
'          sed50 = CalcPercentile(SedResults, 50)
'          sed75 = CalcPercentile(SedResults, 75)
          
          row = rStart + k
          With Sheets("Summary Results - Year 1")
               .Cells(row, colStart) = CStr(Round(sed_avg, 2))
               .Cells(row, colStart + 1) = hs_min
               .Cells(row, colStart + 2) = CStr(Round(sed_min, 2))
               .Cells(row, colStart + 3) = hs_max
               .Cells(row, colStart + 4) = CStr(Round(sed_max, 2))
'               .Cells(row, colStart + 5) = CStr(Round(sed25, 2))
'               .Cells(row, colStart + 6) = CStr(Round(sed50, 2))
'               .Cells(row, colStart + 7) = CStr(Round(sed75, 2))
          End With
          DoEvents
     Next k
Next i

End Sub
Sub GetRunValues(ID As Integer)

'**************************************************************************************
'get peak 10 minute rainfall intensity in mm/hr and js_sedconv values from returned ERMIT results
'**************************************************************************************

Dim i1 As Long
Dim i2 As Long
Dim val As String
Dim rint As Integer
Dim hs_row As Integer

SubName = "GetRunValues"
hs_row = HillslopeStartRow + RunStartRow + ID - 1

rint = 6 + ID

'find peak rainfall intensity
i1 = InStr(1, ErmitResults, "name=" & Chr(34) & "i10" & Chr(34))
If i1 > 0 Then
     i1 = i1 + 9
     i2 = InStr(i1, ErmitResults, ">")
     val = Mid(ErmitResults, i1, i2 - i1)
     val = Replace(val, "value=", "")
     val = Replace(val, Chr(34), "")
     val = Trim(val)
     Sheets("StoredRunValues").Cells(rint, 1) = hs_row
     Sheets("StoredRunValues").Cells(rint, 2) = val
End If

'find js_sedconv value

i1 = InStr(1, ErmitResults, "js_sedconv = ")
If i1 > 0 Then
     i1 = i1 + 12
     i2 = InStr(i1, ErmitResults, Chr(10))
     val = Mid(ErmitResults, i1, i2 - i1)
     Sheets("StoredRunValues").Cells(rint, 3) = Trim(val)
End If
End Sub
Sub SetUnits()

Select Case Sheets("Inputs").cboUnits
     Case "English"
          LenUnits = "ft"
          AreaUnits = "acres"
          
     Case "Metric"
          LenUnits = "m"
          AreaUnits = "hectares"
End Select

End Sub



Sub GetArrays(ID As Long)

'************************************************************************************************************
'Parse out sediment deliver and cumulative probability arrays from ERMIT results
'************************************************************************************************************

Dim results As String
Dim length As Long
Dim l As Long
Dim i1 As Long
Dim i2 As Long
Dim s As String
Dim SedConv As Double
Dim j As Integer
Dim k As Integer
Dim sFind As String
Dim sedval As Single
Dim Line() As String
Dim prob() As String
Dim p() As String
Dim colCount As Integer
Dim colSed As Integer
Dim colCP As Integer
Dim blnEnd As Boolean
Dim sed() As Variant            'sediment delivery array each element is for the model year
Dim cp() As Variant          'cumulative probability array - each row is a treatment, each col is the probability value for a given model year
Dim cp_col As Integer
Dim rPut As Long
Dim yr As Integer
Dim colPut As Integer
Dim plen As Integer
Dim sedlen As Integer
Dim rStart As Long
Dim problen As Long

SubName = "GetArrays"

results = ErmitResults

l = InStr(1, results, "function whatseds")

results = Left(results, l - 1)

'find array count
l = InStr(1, results, "var a_len=")
sedlen = Mid(results, l + 10, 3) - 1

'ReDim sed(sedlen)
'plen = sedlen * 5 - 1
ReDim cp(5, 1000)

'find sediment conversion factor
i1 = InStr(1, results, "js_sedconv = ") + 12
i2 = InStr(1, results, "js_storage_units")
s = Mid(results, i1, i2 - i1)
SedConv = ConvertStringToNumber(s)

i1 = InStr(1, results, "sed_del[1]") - 1

results = Right(results, Len(results) - i1)

j = 1

'find each sediment delivery and cumulative probability array value
Do Until blnEnd = True
     sFind = "sed_del[" & j & "]"
     i1 = InStr(1, results, sFind)
     
     sFind = "sed_del[" & j + 1 & "]"
     i2 = InStr(1, results, sFind)
          
     If i2 = 0 Then           'last array index reached
          i2 = Len(results)
          blnEnd = True
     End If
     
     'parse sediment delivery value out
     s = Mid(results, i1, i2 - i1)
     
     i1 = InStr(1, s, "=") + 1
     i2 = InStr(1, s, ";")
     sedval = Trim(Mid(s, i1, i2 - i1))
     
     sedlen = j - 1
      ReDim Preserve sed(sedlen)
     sed(j - 1) = sedval
     
     i1 = InStrRev(s, "'") + 1
     s = Right(s, Len(s) - i1)

     'parse out cumulative probability values
     Line = Split(s, Chr(10))
     
     'loop through treatments 0-untreated, 1-seeded, 2-mulch 47%, 3-5 mulch
     For k = 0 To UBound(Line)
          prob = Split(Line(k), ";")         'probability values separated by semi-colon
          
          'put values into array
          For l = 0 To UBound(prob)
               p = Split(prob(l), "=")
               cp(k, cp_col + l) = Trim(p(1))
          Next l
     Next k
     cp_col = cp_col + 5
     j = j + 1
Loop

'put array values into worksheets
With Sheets("Sediment")
     rPut = ID + SedStartRow - 1
     
     For j = 0 To UBound(sed)
            .Cells(rPut, j + 1) = CStr(Round(sed(j), 1))
     Next j
End With

rPut = ProbStartRow

problen = (sedlen + 1) * 5 - 1
With Sheets("Probability")
     Do Until .Cells(rPut, 1) = ""
          rPut = rPut + 1
     Loop
     
     For j = 0 To 5      ' j indicates treatment type
          
          For yr = 0 To 4
               colPut = 4
               .Cells(rPut, 1) = ID
               .Cells(rPut, 2) = j
               .Cells(rPut, 3) = yr + 1
               
               For k = yr To problen Step 5
                        .Cells(rPut, colPut) = cp(j, k)
                        colPut = colPut + 1
               Next k
               rPut = rPut + 1
          Next yr
     Next j
End With

End Sub



Sub ClearSheets()

'**********************************************************
'Clear result sheets
'**********************************************************

SubName = "ClearSheets"
Application.ScreenUpdating = False

'clear sediment delivery sheet
With Sheets("Sediment")
     .Activate
          Range("A" & SedStartRow & ":IV65000").Select
          Selection.ClearContents
End With

'clear probability sheet
With Sheets("Probability")
     .Activate
          Range("A" & ProbStartRow & ":IV65000").Select
          Selection.ClearContents
End With

'clear rainfall stats
With Sheets("Results - Rainfall")
.Activate
     Range("A13:H65000").Select
     Selection.ClearContents
End With

'clear stored run values sheet
With Sheets("StoredRunValues")
.Activate
     Range("A7:C5000").Select
     Selection.ClearContents
End With

'clear erosion barrier
With Sheets("Results - Erosion Barriers")
.Activate
      .[_ErosionInputs] = ""
      
     Range("A" & LogStartRow & ":IV65000").Select
     Selection.ClearContents
     .Cells(LogStartRow, 1).Select
End With

'clear year 1 summary results

With Sheets("Summary Results - Year 1")
    .[TotalArea] = ""
     .[_table1] = ""
     .[_table2] = ""
     .[_table3] = ""
     .[_table4] = ""
End With

'clear year 1 by hillslope
With Sheets("Results By Hillslope - Year 1")
     .[_y1results] = ""
End With

'clear out year results
With Sheets("Results - Out Years")
    .Activate
     .[_OutYearInputs] = ""
     Range("A" & OutStartRow & ":IV65000").Select
     Selection.ClearContents
     .Cells(OutStartRow, 1).Select
End With

'clear main results
With Sheets("Results - Treatment")
    .[_TreatmentInputs] = ""
    .[_sedavg1] = ""
    .[_sedtot1] = ""
    .[_sedavg2] = ""
    .[_sedtot2] = ""
    .[_sedavg3] = ""
    .[_sedtot3] = ""
    .[_sedavg4] = ""
    .[_sedtot4] = ""
    .[_sedavg5] = ""
    .[_sedtot5] = ""
    .[_sedavg6] = ""
    .[_sedtot6] = ""
    
     .Activate
     Range("A" & ResultStartRow & ":AD65000").Select
     Selection.ClearContents
     .Cells(ResultStartRow, 1).Select
End With

'clear match types sheet
With Sheets("Match Types")
    .[_SoilTypes] = ""
    .[_TreatmentTypes] = ""
End With

Sheets("Inputs").Activate

Application.ScreenUpdating = True
End Sub


Sub GetRainfallStats(ID As Integer)

'**********************************************************
'Parse rainfall event rankings from ERMIT Results
'**********************************************************

Dim results As String
Dim i1 As Long
Dim i2 As Long
Dim sFind As String
Dim f As String
Dim TableRows() As String
Dim TableCells() As String
Dim sReplace As String
Dim j As Integer
Dim k As Integer
Dim celltext As String
Dim stats As StormStat
Dim rStart As Long
Dim r As Long
Dim code As String
Dim rHS As Long

rHS = HillslopeStartRow + RunStartRow + ID - 1
code = Sheets("Inputs").Cells(rHS, 1)

SubName = "GetRainfallStats"
results = ErmitResults

sFind = "Rainfall Event Rankings and Characteristics from the Selected Storms"

'find rainfall event table
i2 = InStr(1, results, sFind)
i1 = InStrRev(results, "<table", i2)
i2 = InStr(i1, results, "</table>")
results = Mid(results, i1, i2 - i1)

'remove unneccesary html code in the table
results = Replace(results, "<b>", "")
results = Replace(results, "<br>", "")
results = Replace(results, "</b>", "")
results = Replace(results, "</font>", "")
results = Replace(results, "<font size=-2>", "")
results = Replace(results, "&nbsp;", " ")
results = Replace(results, " align=right", "")
results = Replace(results, "</a>", " ")
results = Replace(results, "th bgcolor='#ccffff'", "td")
results = Replace(results, "</th>", "</td>")
results = Replace(results, "<sup>", "")
results = Replace(results, "</sup>", "")
results = Replace(results, "<sub>", "")
results = Replace(results, "</sub>", "")

i1 = InStr(1, results, "<!--")
Do Until i1 = 0
     i2 = InStr(i1, results, Chr(10))
     sReplace = Mid(results, i1, i2 - i1)
     results = Replace(results, sReplace, "")
     i1 = InStr(1, results, "<!--")
Loop

i1 = InStr(1, results, "<a onMouseOver")
Do Until i1 = 0
     i2 = InStr(i1, results, ">")
     sReplace = Mid(results, i1, i2 - i1 + 1)
     results = Replace(results, sReplace, "")
     i1 = InStr(1, results, "<a onMouseOver")
Loop

results = Replace(results, Chr(10), "")

'split table rows into an array
TableRows = Split(results, "</tr>")

'loop through table rows array and find the table cell values for each row
For j = 2 To UBound(TableRows)
     
     TableCells = Split(TableRows(j), "</td>")
     
     If UBound(TableCells) > 5 Then
          For k = 0 To UBound(TableCells)
               celltext = TableCells(k)
               i1 = InStr(1, celltext, "<td>") + 3
               
               If i1 > 0 Then
                    celltext = Trim(Right(celltext, Len(celltext) - i1))
                    
                    Select Case k
                         Case 0
                              stats.Rank = celltext
                         Case 1
                              stats.Runoff = celltext
                         Case 2
                              stats.Precipitation = celltext
                         Case 3
                              stats.Duration = celltext
                         Case 4
                              stats.Peak10 = celltext
                         Case 5
                              stats.Peak30 = celltext
                         Case 6
                              celltext = Replace(celltext, "year", " year")
                              stats.StormDate = celltext
                    End Select
               End If
          Next k
          
          RainStats(j - 2, 0) = stats.Rank
          RainStats(j - 2, 1) = stats.Runoff
          RainStats(j - 2, 2) = stats.Precipitation
          RainStats(j - 2, 3) = stats.Duration
          RainStats(j - 2, 4) = stats.Peak10
          RainStats(j - 2, 5) = stats.Peak30
          RainStats(j - 2, 6) = stats.StormDate
     End If
Next j

'put stats into rainfall sheet
r = 13

With Sheets("Results - Rainfall")
     For j = 0 To 5
            If RainStats(j, 0) <> "" Then
                  .Cells(r, 1) = code
               
                  For k = 0 To 6
                        .Cells(r, k + 2) = RainStats(j, k)
                  Next k
                  r = r + 1
            End If
     Next j
End With

End Sub
Sub ImportHillslopeFile()

'*****************************************************************
'import hillslope data from csv file created by GIS toolbox
'hillslope file format
'DEH made more robust to differing columns possible from Toolbox 2011.11.11
'*****************************************************************

Dim row As Long
Dim col As Integer
Dim i As Long
Dim j As Integer
Dim rStart As Long
Dim sline As String
Dim data() As String
Dim FileData() As String
Dim rowcount As Long
Dim HillSlopeData(13) As Variant
Dim importcount As Integer
Dim sName As String
Dim Area As Single
Dim length1 As Single
Dim length2 As Single
Dim lengthtot As Single
Dim BurnSeverity As Integer
Dim BurnClass As String
Dim BurnClass2 As String
Dim Unburned As Integer

'On Error GoTo EH

   sName = "Inputs"

' BurnClass = Sheets("Inputs").cmbBurnSeverity.Value

   If blnOverwrite = True Then
     ClearHillslopeData
     rStart = HillslopeStartRow
   Else
      rStart = HillslopeStartRow + Sheets("Ref").[InputCount]
   End If

   Open HillslopeFile For Input As #1

   Line Input #1, sline   ' header row

   sline = Replace(sline, Chr(34), "")
   headers = Split(sline, ",")
   
   For i = 0 To UBound(headers)
     If (headers(i) = "Rowid_1") Then Rowid_1 = i   '
     If (headers(i) = "ROWID_") Then ROWID_ = i     '
     If (headers(i) = "OID_") Then OID_ = i         '
     If (headers(i) = "HS_ID") Then hs_ID = i       '
     If (headers(i) = "UNIT_ID") Then UNIT_ID = i   '
     If (headers(i) = "SOIL_TYPE") Then SOIL_TYPE = i ' - GIS land type
     If (headers(i) = "AREA") Then Area_ = i        ' - area
     If (headers(i) = "UTREAT") Then utreat = i     ' - upper treatment
     If (headers(i) = "USLP_LNG") Then USLP_LNG = i ' - upper slope length
     If (headers(i) = "UGRD_TP") Then UGRD_TP = i   ' - upper top gradient %
     If (headers(i) = "UGRD_BTM") Then UGRD_BTM = i ' - upper bottom gradient %
     If (headers(i) = "LTREAT") Then LTREAT = i     ' - lower treatment
     If (headers(i) = "LSLP_LNG") Then LSLP_LNG = i ' - lower slope length
     If (headers(i) = "LGRD_TP") Then LGRD_TP = i   ' - lower section top gradient %
     If (headers(i) = "LGRD_BTM") Then LGRD_BTM = i ' - lower section bottom gradient %
    'If (headers(i) = "ADJ_STRM") Then ADJ_STRM = i ' - not used by WEPP
    'if (headers(i) = "TRIB_TO") then TRIB_TO = i   ' - not used by WEPP
     If (headers(i) = "ERM_TSLP") Then ERM_TSLP = i ' top slope used for ERMIT
     If (headers(i) = "ERM_MSLP") Then ERM_MSLP = i ' - mid slope used for ERMIT
     If (headers(i) = "ERM_BSLP") Then ERM_BSLP = i ' - bottom slope used for ERMiT
     If (headers(i) = "BURNSEV") Then BurnSev = i   ' - burn severity rating
     If (headers(i) = "BURNCLASS") Then BurnClass2 = i   ' - burn severity rating
   Next

'put each row of text into array element
   Do Until EOF(1)
      Line Input #1, sline
      ReDim Preserve FileData(rowcount)
      'remove quotation marks
      sline = Replace(sline, Chr(34), "")
      FileData(rowcount) = sline
      rowcount = rowcount + 1
   Loop
   Close #1

   row = rStart

'put hillslope data into spreadsheet
   With Sheets(sName)
     .Activate
     For i = 0 To UBound(FileData)
        data = Split(FileData(i), ",")
           
'        If UBound(data) >= 19 Then
           If Trim(data(UNIT_ID)) = "" Then
              .Cells(row, 1) = data(hs_ID)
           Else
              .Cells(row, 1) = data(hs_ID) & " _ " & data(UNIT_ID) ' hillslope code combination of HS_ID and UNIT_ID
           End If
           If IsNumeric(data(Area_)) Then                    'area
              Area = data(Area_)
           End If
           If IsNumeric(data(USLP_LNG)) Then                 'upper slope length
              length1 = data(USLP_LNG)
           Else
              length1 = 0
           End If
           If IsNumeric(data(LSLP_LNG)) Then               'lower slope length
              length2 = data(LSLP_LNG)
           Else
              length2 = 0
           End If
           lengthtot = length1 + length2
           If blnConvertHillslopeUnits = True Then
              Select Case .cboUnits
                 Case "English"
                     lengthtot = lengthtot * 3.2808399         'convert meters to feet
                     Area = Area * 2.47105381                        'convert hectares to acres
                  Case "Metric"
                     lengthtot = lengthtot / 3.2808399         'convert feet to meters
                     Area = Area / 2.47105381                        'convert acres to hectares
              End Select
           End If
           .Cells(row, 3) = data(SOIL_TYPE)           'GIS soil type
           .Cells(row, 4) = CStr(Round(Area, 3))          'area
           .Cells(row, 5) = data(utreat)           'GIS treatment
           .Cells(row, 15) = lengthtot         'horizontal length
           .Cells(row, 9) = Round(data(ERM_TSLP), 2)      'upper slope
           .Cells(row, 10) = Round(data(ERM_MSLP), 2)      'mid slope
           .Cells(row, 11) = Round(data(ERM_BSLP), 2)      'toe slope
           If (Not IsEmpty(BurnSev)) Then
               .Cells(row, 16) = data(BurnSev)         'burn severity rating



           End If

           'set burn severity class
           If (Not Len(Trim(BurnClass2)) = 0) Then
               .Cells(row, 17) = data(BurnClass2)
           Else
               BurnSeverity = data(BurnSev)
    
               Select Case BurnSeverity
                  Case Is > .Cells(19, 17)
                     .Cells(row, 17) = "High"
                  Case .Cells(18, 17) + 1 To .Cells(19, 17)
                     .Cells(row, 17) = "Moderate"
                  Case .Cells(17, 17) + 1 To .Cells(18, 17)
                     .Cells(row, 17) = "Low"
                  Case Is <= .Cells(17, 17)
    '                Unburned = Unburned + 1            ' DEH 2014.04.08
    '               .Cells(row, 17) = ""                ' DEH 2014.04.08
                    .Cells(row, 17) = "Unburned"        ' DEH 2014.04.08
               End Select
           End If
  
           importcount = importcount + 1
           row = row + 1
 '       End If
     Next i
     .Activate
     .Cells(HillslopeStartRow, 1).Select
   End With

'   DEH 2014.04.08
'   If Unburned > 0 Then
'      MsgBox (Unburned & " hillslope rows are classified as unburned.  ERMiT currently cannot process unburned hillslopes." & vbCrLf _
'             & "Use Disturbed WEPP to evaluate unburned hillslopes.")
'   End If
   
   GetSoilTypes
   GetTreatmentTypes
   MsgBox importcount & " rows imported"

Exit Sub

EH:
     MsgBox "The following error occurred while importing Hillslope data row: " & i & vbCrLf & Err.Number & ": " & Err.Description
End Sub
Sub ClearHillslopeData()

With Sheets("Match Types")
      .Activate
      .Range("_SoilTypes").Select
      Selection.ClearContents
      .Range("_TreatmentTypes").Select
      Selection.ClearContents
End With

With Sheets("Inputs")
      .Activate
      Range("A" & HillslopeStartRow & ":Q60000").Select
      Selection.ClearContents
      Selection.ClearComments
      Selection.Interior.ColorIndex = 0
      Range("A" & HillslopeStartRow).Select
End With

End Sub
Sub SelectHillslopeFile()

'***************************************************************************
'prompt user to select the hillslope csv file created by the GIS toolbox
'***************************************************************************

Dim f As String
Dim bln As Boolean
Dim FileCheck As String
Dim r As Integer
Dim count As Integer
Dim iResp As Integer
Dim sMessage As String
Dim units As String
Dim FileUnits As String

With Application.FileDialog(msoFileDialogFilePicker)
            .AllowMultiSelect = False
            .Title = "Select hillslope csv file"
            .Filters.Clear
            .Filters.Add "csv files", "*.csv"
            .Show
            If .SelectedItems.count > 0 Then
                  f = .SelectedItems(1)
            End If
End With

If f = "" Then
     Exit Sub
End If

'check if there is already data in the hillslope sheet and give the user the open of clearing or appending
If Sheets("Ref").[InputCount] > 0 Then
     iResp = MsgBox("Do you wish to over write the existing hillslope data in the spreadsheet?" & vbCrLf _
                         & " Select Yes to overwrite and No to append the data.", vbYesNo)
    If iResp = vbYes Then
      blnOverwrite = True
    Else
      blnoverwite = False
   End If
End If

HillslopeFile = f
units = Sheets("Inputs").cboUnits
Select Case units
      Case "English"
            FileUnits = "metric"
      Case "Metric"
            FileUnits = "english"
End Select

sMessage = units & " units are currently selected. " _
            & "  If the hillslope file is in " & FileUnits & " units, select Convert Hillslope File Units" _
            & " otherwise select Import Without Conversion."
With frmImport
      .Label1.Caption = sMessage
      .Show
End With



End Sub
Sub SetClimateRows()

'for multiple climate option initially set climate rows to be the first climate
Dim firstclimate As String
Dim rCell As Range

For Each rCell In Sheets("Ref").[_climates]
     firstclimate = rCell.Value
     Exit For
Next rCell

With Sheets("Inputs")
     For r = HillslopeStartRow To 500
          .Cells(r, 2) = ""
     Next r
End With

End Sub

Sub UnhideAllSheets()

Dim sheet As Object

For Each sheet In Sheets
      sheet.Visible = xlSheetVisible
Next sheet

Set sheet = Nothing
End Sub
Sub LogErmitRun(count, climate)

'=================== DEH ====================================
' send data to perl program weblogger.pl to log the batch run
'============================================================

SubName = "LogErmitRun"

Dim strPost As String
Dim results As String
Dim xmlhttp As Object
Dim Version As String

Version = "2018.04.01"

strPost = "count=" & count & "&version=" & Version & "&climate=" & climate

'replace spaces with underscore
strPost = Replace(strPost, " ", "_")

'remove commas
strPost = Replace(strPost, ",", "")

Set xmlhttp = CreateObject("MSXML2.XMLHTTP")

xmlhttp.Open "POST", "https://forest.moscowfsl.wsu.edu/cgi-bin/fswepp/ermit/weblogger.pl ", False
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xmlhttp.send (strPost)
results = xmlhttp.responseText

Set xmlhttp = Nothing

End Sub

Sub GetTreatmentTypes()

'lists each unique treatment type from the GIS Hillslope file in the Match Types sheet, so
'the user can match to the ERMit vegetationType

Dim row As Long
Dim Types() As String
Dim Treatment As String
Dim LastRow As Long
Dim count As Integer
Dim blnMatch As Boolean
Dim i As Integer
Dim j As Integer
Dim HSCount As Integer
Dim r1 As Integer
Dim c1 As Integer
Dim sRange As String
Dim a() As String

'get start row of soil types range in match types sheet
sRange = ActiveWorkbook.Names("_TreatmentTypes").RefersToR1C1
a = GetRangeCoordinates(sRange)
r1 = a(0)
c1 = a(1)

HSCount = Sheets("Ref").Range("InputCount")
LastRow = HillslopeStartRow + HSCount - 1

'loop through hillslope data and get each uniqute upper treatment type
For row = HillslopeStartRow To LastRow
            Treatment = Trim(Sheets("Inputs").Cells(row, 5))
            
          If count = 0 Then
                  ReDim Types(0)
                  Types(0) = Treatment
                  count = count + 1
            Else
                  blnMatch = False
                  
                  'determine if treatment type already is in array
                  For j = 0 To UBound(Types)
                      If Treatment = Types(j) Then
                          blnMatch = True
                          Exit For
                        End If
                  Next j
            
                  If Not blnMatch Then
                        ReDim Preserve Types(count)
                        Types(count) = Treatment
                        count = count + 1
                  End If
            End If
Next row
      
'put treatment t ypes into Match  Types Sheet
With Sheets("Match Types")
            .Activate
      .Range("_TreatmentTypes").Select
      Selection.ClearContents
      
      For i = 0 To UBound(Types)
            .Cells(r1 + i, c1) = Types(i)
      Next i
      
End With
End Sub
Sub GetSoilTypes()

'lists each unique soil type from the GIS Hillslope file in the Match  Types sheet, so
'the user can match to the disturbed ERMiT Soil Type

Dim row As Long
Dim Types() As String
Dim SoilType As String
Dim LastRow As Long
Dim count As Integer
Dim blnMatch As Boolean
Dim i As Integer
Dim j As Integer
Dim HSCount As Integer
Dim r1 As Integer
Dim c1 As Integer
Dim sRange As String
Dim a() As String

'get start row of soil types range in match types sheet
sRange = ActiveWorkbook.Names("_SoilTypes").RefersToR1C1
a = GetRangeCoordinates(sRange)
r1 = a(0)
c1 = a(1)

HSCount = Sheets("Ref").Range("InputCount")

'loop through hillslope data and get each uniqute soil type
LastRow = HillslopeStartRow + HSCount - 1

For row = HillslopeStartRow To LastRow
            SoilType = Trim(Sheets("Inputs").Cells(row, 3))
            
          If count = 0 Then
                  ReDim Types(0)
                  Types(0) = SoilType
                  count = count + 1
            Else
                  blnMatch = False
                  
                  'determine if soil type already is in array
                  For j = 0 To UBound(Types)
                      If SoilType = Types(j) Then
                          blnMatch = True
                          Exit For
                        End If
                  Next j
            
                  If Not blnMatch Then
                        ReDim Preserve Types(count)
                        Types(count) = SoilType
                        count = count + 1
                  End If
            End If
Next row
      
'put soil t ypes into Match Soil Types Sheet
With Sheets("Match Types")
      .Activate
      .Range("_SoilTypes").Select
      Selection.ClearContents
      
      For i = 0 To UBound(Types)
            .Cells(r1 + i, c1) = Types(i)
      Next i
      
      .Cells(r1, c1).Select
End With
      
End Sub
Public Function GetRangeCoordinates(ByVal sRange As String) As Variant

'converts string range to comma separated string of numeric cell coordinates r1,c1,r2,c2
Dim i1 As Integer
Dim s As String
Dim a() As String

i1 = InStr(1, sRange, "!")

s = Right(sRange, Len(sRange) - i1)
s = Replace(s, ":", "")
s = Replace(s, "R", ",")
s = Replace(s, "C", ",")
s = Right(s, Len(s) - 1)

a = Split(s, ",")

GetRangeCoordinates = a

End Function
Sub UpdateType(ByVal sType As String, ByVal GIS_Type As String, ByVal Ermit_Type As String)

Dim col1 As Integer
Dim col2 As Integer
Dim row1 As Long
Dim row2 As Long
Dim count As Long
Dim s1 As String

count = Sheets("Ref").Range("InputCount")

'update ERMiTType
Select Case sType
      Case "Soil"
            col1 = 3
            col2 = 6
            
      Case "Vegetation"
            col1 = 5
            col2 = 8
      
End Select

row1 = HillslopeStartRow
row2 = HillslopeStartRow + count - 1

'loop through hillslope data and update matched rows
For row = row1 To row2
      s1 = Sheets("Inputs").Cells(row, col1)
      
      If s1 = GIS_Type Then
            Sheets("Inputs").Cells(row, col2) = Ermit_Type
      End If
Next row
End Sub

Sub ExportBurnSeverity()

'export hillslope ID and burn severity rating to a csv file for import back into GIS

Dim f As String
Dim r As Long
Dim count As Long
Dim rStop As Long

count = Sheets("Ref").Range("InputCount")
rStop = HillslopeStartRow + count - 1

f = Application.GetSaveAsFilename("Burn_Rating.csv", "CSV Files (*.csv), *.csv", , "Export Burn Severity Rating To CSV File")

Open f For Output As #1

Print #1, "HS_ID, Burn_Severity"

For r = HillslopeStartRow To rStop
    Print #1, Sheets("Inputs").Cells(r, 1) & "," & Sheets("Inputs").Cells(r, 17)
Next r

Close #1
End Sub


