
'**********************************************************
'CONSTANTS
'**********************************************************
Public Const HillslopeStartRow = 25    'row where hillslope input data starts
Public Const SedStartRow = 3                 'row where sediment delivery values start
Public Const ProbStartRow = 5                'row where probability values start
Public Const LogStartRow = 18                'row where log results start
Public Const OutStartRow = 18                'row where out year results start
Public Const ResultStartRow = 24
Public Const ClimateStartRow = 71

'constants used in logs and wattles calculation
Public Const coeff_slope = 1342               ' slope in whole percent i.e. '30' used in log calculation
Public Const coeff_diam = 0.0029             '  diam^2 for diam in cm
Public Const coeff_spacing = 272             ' spacing in m
Public Const Intercept = -35.4

'**********************************************************
'PUBLIC VARIABLES
'**********************************************************
Public InputErrorCount As Long
Public LenUnits As String
Public AreaUnits As String
Public ErmitResults As String
Public RainStats(5, 6) As String
Public UnitSystem As String
Public SubName As String
Public blnRunning As Boolean
Public blnStopRun As Boolean
Public blnOverwrite As Boolean
Public HillslopeFile As String
Public blnConvertHillslopeUnits As Boolean
Public CurrentSheet As Integer


'**********************************************************
'TYPES
'**********************************************************
Type ErmitInput
      hs_code As String
      Area As String
     SoilType As String
     VegType As String
     RockPercent As Single
     TopGradient As Single
     MidGradient As Single
     ToeGradient As Single
     ShrubPercent As String
     GrassPercent As String
     BarePercent As String
     HorizontalSlopeLength As Single
     BurnClass As String
End Type

Type StormStat
     Rank As String
     Runoff As String
     Precipitation As String
     Duration As String
     Peak10 As String
     Peak30 As String
     StormDate As String
End Type
