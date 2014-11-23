Option Explicit

Sub OpenFiles()
Application.ScreenUpdating = False 'change to False    --- False
Application.Calculation = xlCalculationManual 'change to xlCalculationManual       xlCalculationAutomatic

  'variables
  Dim wbResults, wbCodeBook As Workbook
  Dim wsSummarySheet As Worksheet
  Dim wbCodeBookName, Column, strPath, strFile, strDir As String
  Dim Record(9, 4, 39) As Variant 'array
  Dim fso, objFiles, obj As Object
  Dim FileCount As Integer
    
  'counters
  Dim i, j, sumj, t, arrCounter, arrCounterA, mmf As Integer    'i is iterator counter, j is [temp]column counter for pasting on DB tab, t is file counter iterator
    
  'initialize variables
  Set wbCodeBook = Application.ActiveWorkbook
  wbCodeBookName = ActiveWorkbook.Name
  If i = 0 Then i = 1

  'Where are you aiming to grab files from?
  strPath = ActiveWorkbook.Worksheets("Run Macro").Range("E12").Value
      'Ensure the path ends in a backslash
      If Right(strPath, 1) <> "\" Then
          strPath = strPath & "\"
      End If
  
  'Call the first file within the folder (change the file extension, accordingly)
  strFile = Dir(strPath & "*.xls*")
  
  'Loop through each file within the folder
  Do While Len(strFile) > 0
      Set wbResults = Workbooks.Open(strPath & strFile, UpdateLinks:=0)
      Windows(strFile).Activate
      Set wsSummarySheet = ActiveWorkbook.Worksheets("Template")
      'Store SummarySheet values in memory via the Record array
      ' Static variables
      Record(0, 0, 1) = wsSummarySheet.Range("b3").Value         'AccountId
      Record(0, 0, 2) = wsSummarySheet.Range("b2").Value         'Company Name
      
      Record(0, 0, 3) = wsSummarySheet.Range("b16").Value         'Contract_Date
      Record(0, 0, 4) = wsSummarySheet.Range("c16").Value        'Contract_Visa_Rate_Variable
      Record(0, 0, 5) = wsSummarySheet.Range("d16").Value        'Contract_Visa_Rate_Fixed
      Record(0, 0, 6) = wsSummarySheet.Range("e16").Value        'Contract_MC_Rate_Variable
      Record(0, 0, 7) = wsSummarySheet.Range("f16").Value        'Contract_MC_Rate_Fixed
      Record(0, 0, 8) = wsSummarySheet.Range("g16").Value        'Contract_Debit_Rate_Variable
      Record(0, 0, 9) = wsSummarySheet.Range("h16").Value        'Contract_Debit_Rate_Fixed
      Record(0, 0, 10) = wsSummarySheet.Range("i16").Value        'Contract_DC_Rate_Variable
      Record(0, 0, 11) = wsSummarySheet.Range("j16").Value        'Contract_DC_Rate_Fixed
      Record(0, 0, 12) = wsSummarySheet.Range("k16").Value        'Contract_eCheck_Rate_Fixed
      Record(0, 0, 13) = wsSummarySheet.Range("l16").Value        'Contract_Amex_Rate_Variable
      Record(0, 0, 14) = wsSummarySheet.Range("m16").Value        'Contract_Scan_Rate_Fixed
      Record(0, 0, 15) = wsSummarySheet.Range("n16").Value        'Contract_MMF_Rate_Variable
      Record(0, 0, 16) = wsSummarySheet.Range("o16").Value        'Contract_PayByCash_Rate_Fixed
      
      Record(0, 0, 17) = wsSummarySheet.Range("b17").Value         'Addendum_Date
      Record(0, 0, 18) = wsSummarySheet.Range("c17").Value        'Addendum_Visa_Rate_Variable
      Record(0, 0, 19) = wsSummarySheet.Range("d17").Value        'Addendum_Visa_Rate_Fixed
      Record(0, 0, 20) = wsSummarySheet.Range("e17").Value        'Addendum_MC_Rate_Variable
      Record(0, 0, 21) = wsSummarySheet.Range("f17").Value        'Addendum_MC_Rate_Fixed
      Record(0, 0, 22) = wsSummarySheet.Range("g17").Value        'Addendum_Debit_Rate_Variable
      Record(0, 0, 23) = wsSummarySheet.Range("h17").Value        'Addendum_Debit_Rate_Fixed
      Record(0, 0, 24) = wsSummarySheet.Range("i17").Value        'Addendum_DC_Rate_Variable
      Record(0, 0, 25) = wsSummarySheet.Range("j17").Value        'Addendum_DC_Rate_Fixed
      Record(0, 0, 26) = wsSummarySheet.Range("k17").Value        'Addendum_eCheck_Rate_Fixed
      Record(0, 0, 27) = wsSummarySheet.Range("l17").Value        'Addendum_Amex_Rate_Fixed
      Record(0, 0, 28) = wsSummarySheet.Range("m17").Value        'Addendum_Scan_Rate_Fixed
      Record(0, 0, 29) = wsSummarySheet.Range("n17").Value        'Addendum_MMF_Rate_Variable
      Record(0, 0, 30) = wsSummarySheet.Range("o17").Value        'Addendum_PayByCash_Rate_Fixed
      
      Record(0, 0, 31) = wsSummarySheet.Range("b18").Value         'Debit_letter
      Record(0, 0, 32) = wsSummarySheet.Range("g18").Value        'Debit_Debit_Rate_Variable
      Record(0, 0, 33) = wsSummarySheet.Range("h18").Value        'Debit_Debit_Rate_Fixed
      
      Record(0, 0, 34) = wsSummarySheet.Range("b19").Value         'Install Date
      Record(0, 0, 35) = wsSummarySheet.Range("b20").Value         'Date Inactivated
      Record(0, 0, 36) = wsSummarySheet.Range("b12").Value         'Error Type
      Record(0, 0, 37) = wsSummarySheet.Range("b13").Value         'Action Item
      Record(0, 0, 38) = wsSummarySheet.Range("e6").Value         'Notes
      
      ' Variable variables
      For arrCounter = 1 To (3)  '2011
        Record(1, arrCounter, 0) = "201" & arrCounter                                    'Year
        Record(2, arrCounter, 0) = wsSummarySheet.Range("b" & arrCounter + 5).Value      'Internal
        Record(3, arrCounter, 0) = wsSummarySheet.Range("c" & arrCounter + 5).Value      'Contractor
      Next
      
      mmf = 0
      For arrCounterA = 1 To 3
        Record(4, arrCounterA, 0) = wsSummarySheet.Range("n" & arrCounterA + 30 + mmf).Value      'mmf
        If mmf < 10 Then
          mmf = mmf + 14 - arrCounterA
        Else
          mmf = mmf + 15 - arrCounterA
        End If
        
      Next
        
      'Continue adding variables here etc etc
      
      Windows(wbCodeBookName).Activate
      Sheets("DB").Select
      'Paste the values to the DB
      For j = 1 To (UBound(Record, 2) - 1) 'ie for A to P, column values -- j is inner loop counter, UBound determines the size of the array
          Range("a" & j + i).Value = Record(0, 0, 1)    ' account id
          Range("b" & j + i).Value = Record(0, 0, 2)    ' name
          Range("c" & j + i).Value = Record(0, 0, 3)    ' contract date
          Range("d" & j + i).Value = Record(0, 0, 4)    ' contract visa rate variable
          Range("e" & j + i).Value = Record(0, 0, 5)    ' contract visa rate fixed
          Range("f" & j + i).Value = Record(0, 0, 6)    ' contract MC rate variable
          Range("g" & j + i).Value = Record(0, 0, 7)    ' contract MC rate fixed
          Range("h" & j + i).Value = Record(0, 0, 8)    ' contract Debit rate variable
          Range("i" & j + i).Value = Record(0, 0, 9)    ' contract Debit rate fixed
          Range("j" & j + i).Value = Record(0, 0, 10)    ' contract DC rate variable
          Range("k" & j + i).Value = Record(0, 0, 11)    ' contract DC rate fixed
          Range("l" & j + i).Value = Record(0, 0, 12)    ' contract eCheck fixed
          Range("m" & j + i).Value = Record(0, 0, 13)    ' contract Amex fixed
          Range("m" & j + i).Value = Record(0, 0, 14)    ' contract Scan fixed
          Range("o" & j + i).Value = Record(0, 0, 15)    ' contract MMF fixed
          Range("p" & j + i).Value = Record(0, 0, 17)    ' contract PayByCash fixed
          
          Range("q" & j + i).Value = Record(0, 0, 17)    ' addendum date
          Range("r" & j + i).Value = Record(0, 0, 18)    ' addendum visa rate variable
          Range("s" & j + i).Value = Record(0, 0, 19)    ' addendum visa rate fixed
          Range("t" & j + i).Value = Record(0, 0, 20)    ' addendum MC rate variable
          Range("u" & j + i).Value = Record(0, 0, 21)    ' addendum MC rate fixed
          Range("v" & j + i).Value = Record(0, 0, 22)    ' addendum Debit rate variable
          Range("w" & j + i).Value = Record(0, 0, 23)    ' addendum Debit rate fixed
          Range("x" & j + i).Value = Record(0, 0, 24)    ' addendum DC rate variable
          Range("y" & j + i).Value = Record(0, 0, 25)    ' addendum DC rate fixed
          Range("z" & j + i).Value = Record(0, 0, 26)    ' addendum eCheck fixed
          Range("aa" & j + i).Value = Record(0, 0, 27)    ' addendum Amex fixed
          Range("ab" & j + i).Value = Record(0, 0, 28)    ' addendum Scan fixed
          Range("ac" & j + i).Value = Record(0, 0, 29)    ' addendum MMF fixed
          Range("ad" & j + i).Value = Record(0, 0, 30)    ' addendum PayByCash fixed
          
          Range("ae" & j + i).Value = Record(0, 0, 31)    ' debit date
          Range("af" & j + i).Value = Record(0, 0, 32)    ' debit debit rate variable
          Range("ag" & j + i).Value = Record(0, 0, 33)    ' debit debit rate fixed
          
          Range("ah" & j + i).Value = Record(0, 0, 34)    ' install date
          Range("ai" & j + i).Value = Record(0, 0, 35)    ' date inactivated
          Range("aj" & j + i).Value = Record(0, 0, 36)    ' error type
          Range("ak" & j + i).Value = Record(0, 0, 37)    ' action item
          Range("al" & j + i).Value = Record(0, 0, 38)    ' notes
          
          Range("am" & j + i).Value = Record(1, j, 0)    ' year
          Range("an" & j + i).Value = Record(2, j, 0)    ' internal
          Range("ao" & j + i).Value = Record(3, j, 0)    ' contractor
          Range("ap" & j + i).Value = Record(4, j, 0)    ' mmf
          'Sheets("DB").Range(Column & i).Value = Record(j)
          'Column = Chr(Asc(Column) + 1)
      Next j

    'Close the looped file without savingCo
        Windows(strFile).Activate
        wsSummarySheet.Select
        wbResults.Close savechanges:=False
        strFile = Dir
    
    'iterator / reset loop variables
    Column = "A"    ' reset column variable
    'increment the counter by one for each time you run the macro
    sumj = UBound(Record, 2)
    i = sumj + i
  Loop
  
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub


