Attribute VB_Name = "Module11211"

Sub DFH_Degradation_V5A()
'
' macro1 Macro
' Macro recorded 2/21/2013 by chung_su

' 2013-03-18 : Modified for headers and test sequence
' 2013-04-16 : INTO ABS  1KOE
' 2013-05-07 : num of group is 17 since DFH stress up to 4V.
' 2013-05-20 : Add col-count /row count funciton
' 2013-07-31 : config, config_wf assigned to wafer
' 2013-08-05   : Flag for high DFH_R, Change "Config" to "wafer"
' 2013-09-09  : REDUCE THE LAST STEP NOT USED FOR DATA ANALYSIS
' 2015-07-06       sman AVG rms
' 2015-07-30  'For PLD V5A,'romove re-PLD tested headers, 'move init in-situ data to the end cols



Dim first_Col, k, j, i, n, num_Group, row, col, a As Integer
Dim row_Step, col_step As Integer
Dim num_of_Row, max_row As Integer
Dim DFH_V As String
Dim Ambient_Temp, Temp_Coeff, Temp_Median As Double
Dim Rng As Range
Dim init_val, bad_DFH_Rate As Double
Dim intResponse As Integer
Dim testName, StressVolt, StressTime, StressField As String
Dim Count, badDFH_R As Integer



badDFH_R = 0
max_row = 3000

first_Col = 32     ' 29        ' first column need to be cut and paste
num_Group = 13     '13
num_of_Row = 0
col_step = 11      '8

'***Need to Confirm each time
Ambient_Temp = 23
Temp_Coeff = 1.5
DFH_R = 105

'*********************************************************************
'**** Temp coeff : LD37=0.8      D60L/FB : 1.68 (RO) &  1.5 (FD)  m34 : 1.49 ****************     2012-02-08
'*********************************************************************

'romove re-PLD tested headers
k = 1
Do
    k = k + 1
    If Left(Cells(k, 1), 2) <> "SR" Then
        Cells(k, 1).EntireRow.Delete
    End If

Loop Until IsEmpty(Cells(k, 1))
     


'move init in-situ data to the end cols
  Columns("AF:AJ").Select
    Selection.Cut
    Range("HU1").Select
    ActiveSheet.Paste
    
    Columns("AF:AJ").Select
    Selection.Delete Shift:=xlToLeft
    
    


k = 1
Do
    k = k + 1
    If Left(Cells(k, 1), 2) <> "SR" Then
        Cells(k, 1).EntireRow.Delete
    End If
Loop Until IsEmpty(Cells(k, 1))



'find number of cells  (excluding header)
Cells(2, 1).Select

Do
    num_of_Row = num_of_Row + 1
    ActiveCell.Offset(1, 0).Select
Loop Until IsEmpty(ActiveCell.Offset(0, 1))




For i = 2 To num_of_Row
   If Cells(i, 10) > 120 Or Cells(i, 10) < 70 Then
        badDFH_R = badDFH_R + 1
   End If
Next i

bad_DFH_Rate = badDFH_R / num_of_Row * 100
bad_DFH_Rate = Format(bad_DFH_Rate, "##.0")


If bad_DFH_Rate > 5 Then
     i = MsgBox("DFH FR (%) : " & Str(bad_DFH_Rate) & "% " & vbLf & "High DFH_R Fail", vbExclamation)
    intResponse = MsgBox("High DFH FR (%)!!!" & vbLf & "DFH FR (%) : " & Str(bad_DFH_Rate) & "%" & vbLf & "Continue?", vbYesNo)
Else
     
     intResponse = MsgBox("DFH FR (%) : " & Str(bad_DFH_Rate) & "%" & vbLf & "Continue?", vbYesNo)
End If



If intResponse = vbNo Then
     GoTo reTry
End If



intResponse = MsgBox("Ambient Temp / Temp Coeff =" & Ambient_Temp & "C  " & Temp_Coeff & "C/mW   DFH_R= " & CStr(DFH_R) & "?", vbYesNo)


If intResponse = vbNo Then
     GoTo reTry
End If





'cut paste
For k = 0 To num_Group         'Group is the number of the DFH voltage Step
        ActiveSheet.Range(Cells(2, first_Col + k * col_step), Cells(1 + num_of_Row, first_Col + (k + 1) * col_step - 1)).Select
        Selection.Cut
        Cells(2 + num_of_Row * (k + 1), first_Col - col_step).Select
     ActiveSheet.Paste
        
'First common columns copy
          ActiveSheet.Range(Cells(2, 1), Cells(1 + num_of_Row, 20)).Select
          Selection.Copy
          Cells(2 + num_of_Row * (k + 1), 1).Select
          ActiveSheet.Paste

Next k


'NP & MNP
   Cells(1, first_Col) = "NP"
   Cells(1, first_Col + 1) = "MNP"
   'find new num of row
   Cells(2, 1).Select
Do
    k = k + 1                                            'k is row num
    ActiveCell.Offset(1, 0).Select
Loop Until IsEmpty(ActiveCell.Offset(1, 0))

For i = 2 To k
If (Not IsEmpty(Cells(i, 25))) And (Not IsEmpty(Cells(i, 26))) Then    ' Reistance value exist
Cells(i, first_Col) = Cells(i, 27) ^ 2 / Cells(i, 25) / 1000                        ' "NP"
Cells(i, first_Col + 1) = Cells(i, 28) ^ 2 / Cells(i, 25) / 1000                                 '"MNP"
End If

Next i





'Temp and dParameter
For k = 0 To num_Group         'Group is the number of the DFH voltage Step
      
    
        'Set Voltage column values
        Cells(1, first_Col + k * col_step + 2).Select
        
     '   TestName
            testName = Left(ActiveCell, 3)


        
        If Left(testName, 1) = "M" Then
            DFH_V = 0
            Else: DFH_V = Val(Left(testName, 3))
        End If
        
       'Volt and temp value
        For j = 1 To num_of_Row
            Cells(1 + num_of_Row * (k + 1) + j, first_Col + 2) = testName                      'name
            Cells(1 + num_of_Row * (k + 1) + j, first_Col + 3) = k + 2                         'seq
            Cells(1 + num_of_Row * (k + 1) + j, first_Col + 4) = Ambient_Temp + Temp_Coeff * Val(DFH_V) ^ 2 / DFH_R * 1000   'temp
            Cells(1 + num_of_Row * (k + 1) + j, first_Col + 4) = CInt(Cells(1 + num_of_Row * (k + 1) + j, first_Col + 4))
            Cells(1 + num_of_Row * (k + 1) + j, first_Col + 9) = CStr(k + 2) & "_" & CStr(Cells(1 + num_of_Row * (k + 1) + j, first_Col + 4)) & "C"
            'Test Sequence name
            If Mid(Cells(1 + num_of_Row * (k + 1) + j, first_Col + 9), 2, 1) = "_" Then
                Cells(1 + num_of_Row * (k + 1) + j, first_Col + 5) = "0" & Cells(1 + num_of_Row * (k + 1) + j, first_Col + 9)
             Else
                Cells(1 + num_of_Row * (k + 1) + j, first_Col + 5) = Cells(1 + num_of_Row * (k + 1) + j, first_Col + 9)
            End If
          
            '###Nominal temp
             Cells(1 + num_of_Row * (k + 1) + j, first_Col + 6) = ""
            Cells(1 + num_of_Row * (k + 1) + j, first_Col + 7) = ""
            '###&&&&& StressVolt
            If Left(testName, 1) = "M" Then
                StressVolt = "0"
                Else: StressVolt = Left(testName, 4)
            End If
            Cells(1 + num_of_Row * (k + 1) + j, first_Col + 8) = StressVolt          'Angle
            
            
             '####  'Stress Time
            If InStr(testName, "1S") Then
                StressTime = "1sec"
            ElseIf InStr(testName, "5S") Then
                StressTime = "5sec"
            Else: StressTime = ""
            End If
            Cells(1 + num_of_Row * (k + 1) + j, first_Col + 9) = StressTime          'stress time
            
            
               '####  'Stress Field
            
           If InStr(testName, "MFI") Then
                StressField = Left(testName, 4)
                
            Else: StressField = "K"
            End If
            Cells(1 + num_of_Row * (k + 1) + j, first_Col + 10) = StressField          'stress Field
                  
  Next j
  Next k
  
  
   Cells(1, first_Col + 2) = "Test Name"
   Cells(1, first_Col + 3) = "Seq"
   Cells(1, first_Col + 4) = "Temp"
   Cells(1, first_Col + 5) = "Test_Sequence"
   Cells(1, first_Col + 6) = "wafer"
   Cells(1, first_Col + 7) = "Config_wf"
   Cells(1, first_Col + 8) = "StressVolt"
   Cells(1, first_Col + 9) = "DUMMY1"
   Cells(1, first_Col + 10) = "DUMMY2"
   Rows("1:1").Select
   Selection.Replace What:="Initial.", Replacement:=""
    
  '   Columns("ag:ag").Select
 '   Selection.Replace What:=".Bark", Replacement:=""
 '   Selection.Replace What:=".Barkh", Replacement:=""
 '   Selection.Replace What:=".Ba", Replacement:=""
   ' Selection.Replace What:=".", Replacement:=""
    

Count = row_Count(1)

For i = 2 To Count     ' Config_wf dummy fill

   Cells(i, first_Col + 6) = Left(Cells(i, 2), 4)
   Cells(i, first_Col + 7) = Left(Cells(i, 2), 4)
   Cells(i, first_Col + 9) = 0
   Cells(i, first_Col + 10) = 0
    
Next i



For i = 2 To num_of_Row + 1

    For j = 2 To 10
        If IsEmpty(Cells(i, first_Col + j)) Then
            Cells(i, first_Col + j) = 0
            Cells(i, first_Col + 2) = "init"
            Cells(i, first_Col + 3) = "1"
            Cells(i, first_Col + 4) = Ambient_Temp
            Cells(i, first_Col + 5) = "0" & Cells(i, first_Col + 3) & "_" & Cells(i, first_Col + 4)
        End If
    Next j
Next i

' below is only from Ver5A
   Columns("AQ:HO").Select
   Application.CutCopyMode = False
   Selection.Delete Shift:=xlToLeft




    MsgBox ("Completed!!!")
  
reTry:
    
    

    
End Sub




Function row_Count(ByVal col_num As Integer) As Integer
'find number of rows
Dim n As Integer
n = 0
    Cells(1, col_num).Select
    Do
        n = n + 1
        ActiveCell.Offset(1, 0).Select
    Loop Until IsEmpty(ActiveCell.Offset(1, 0))
    row_Count = n + 1
End Function

Function Col_Count(ByVal row_num As Integer) As Integer
'find number of rows
Dim n As Integer
n = 0
    Cells(row_num, 1).Select
    Do
        n = n + 1
        ActiveCell.Offset(0, 1).Select
    Loop Until IsEmpty(ActiveCell.Offset(0, 1))
    Col_Count = n + 1
End Function









