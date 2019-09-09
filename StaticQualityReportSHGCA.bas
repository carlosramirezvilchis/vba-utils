Attribute VB_Name = "StaticQualityReportSHGCA"

Sub Static_Quality_Report_GCA()
Dim hola as String
Dim Level2GCA As String
Dim Level3GCA As String
Dim Level4GCA As String
Dim Level5GCA As String
Dim FaultGCA As String
Dim vCountOfRowsGCA As Integer
Dim vCountOfRowsGCA2 As Integer

'==============================================================================================================================================
'========================================================== GCA C1UL CADILAC ==================================================================
'==============================================================================================================================================
    

FileNameGCAN7 = VBA.Dir("C:\Users\SZ5CM6\Documents\SPHL\GCA\GCAParetoN7.xlsx")

If FileNameGCAN7 = "" Then ' Condition to evaluate if the file exists
MsgBox ("File GCAParetoN7.xlsx does no exists." & vbLf & "Click Ok to continue")
GoTo second

Else
' Opens the GCA NE File in the designated path
Workbooks.Open "C:\Users\SZ5CM6\Documents\SPHL\GCA\" & FileNameGCAN7

Windows("GCAParetoN7.xlsx").Activate

' Returns the last row before the cleanse of the data
vCountOfRowsGCA = Range("A1", Range("A1").End(xlDown)).Rows.Count

' Cycle to only keep "Fits - Exterior"
For a = 6 To vCountOfRowsGCA Step 1
    If Cells(a, 3).Value <> "Fits - Exterior" Then
    
     Rows(a).EntireRow.Clear
    
    End If
Next a

' Orders in descending order, i.e. by WDPV
Range("A5:T10000").Sort Key1:=Range("S6"), Order1:=xlDescending, Header:=xlYes

'============================================= Information transfer for GCA C1UL CADILAC ==========================================================
' Cycle to initialize the variables for each level and wdpv, and to write each variable in the Report file
For i = 6 To 8 Step 1
 
    Windows("GCAParetoN7.xlsx").Activate
 
    Level2GCA = Cells(i, 5).Value
    Level3GCA = Cells(i, 7).Value
    Level4GCA = Cells(i, 9).Value
    Level5GCA = Cells(i, 11).Value
    FaultGCA = Cells(i, 13).Value
    WDPVGCA = Cells(i, 19).Value
    
    Windows("SH Static quality report.xlsm").Activate
    Sheets("Metrics").Select
    
    
    ' Writes the variables in the report file, here, the i variable must change from metric to metric (GCA,DRR,DRL,etc.)
    Cells(i, 6).Value = Level2GCA & "-" & Level3GCA & "-" & Level4GCA & "-" & Level5GCA & "-" & FaultGCA
    Cells(i, 11).Value = WDPVGCA
    
   
Next i


'============================================= Screenshot for GCA C1UL CADILAC ==================================================
Windows("GCAParetoN7.xlsx").Activate
' Returns the last row after the cleanse of the data
vCountOfRowsGCA2 = Range("A1", Range("A1").End(xlDown)).Rows.Count

Range(Cells(5, 1), Cells(12, 19)).Select

Application.CutCopyMode = False
Selection.Copy
Selection.CopyPicture xlScreen, xlBitmap


' Pastes the picture in the corresponding range
Windows("SH Static quality report.xlsm").Activate
Worksheets("GCA").Select
Range("D7").Select
Selection.PasteSpecial
' Changes size of image
    With Selection
                    .ShapeRange.LockAspectRatio = msoFalse
                    .ShapeRange.Height = 200
                    .ShapeRange.Width = 1065
    End With

Windows("GCAParetoN7.xlsx").Close SaveChanges:=True
End If
'==============================================================================================================================================
'========================================================== GCA C1UTL CADILAC =================================================================
'==============================================================================================================================================
second:


FileNameGCANE = VBA.Dir("C:\Users\SZ5CM6\Documents\SPHL\GCA\GCAParetoNE.xlsx")

If FileNameGCANE = "" Then ' Condition to evaluate if the file exists
MsgBox ("File GCAParetoNE.xlsx does no exists." & vbLf & "Click Ok to continue")
GoTo third

Else
 ' Opens the GCA NE File in the designated path
Workbooks.Open "C:\Users\SZ5CM6\Documents\SPHL\GCA\" & FileNameGCANE

Windows("GCAParetoNE.xlsx").Activate

' Returns the last row before the cleanse of the data
vCountOfRowsGCA = Range("A1", Range("A1").End(xlDown)).Rows.Count

' Cycle to only keep "Fits - Exterior"
For a = 6 To vCountOfRowsGCA Step 1
    If Cells(a, 3).Value <> "Fits - Exterior" Then
    
     Rows(a).EntireRow.Clear
    
    End If
Next a

' Orders in descending order, i.e. by WDPV
Range("A5:T10000").Sort Key1:=Range("S6"), Order1:=xlDescending, Header:=xlYes

'============================================= Information transfer for GCA C1UTL CADILAC ==========================================================

' Cycle to initialize the variables for each level and wdpv, and to write each variable in the Report file
For i = 6 To 8 Step 1
 
    Windows("GCAParetoNE.xlsx").Activate
 
    Level2GCA = Cells(i, 5).Value
    Level3GCA = Cells(i, 7).Value
    Level4GCA = Cells(i, 9).Value
    Level5GCA = Cells(i, 11).Value
    FaultGCA = Cells(i, 13).Value
    WDPVGCA = Cells(i, 19).Value
    
    Windows("SH Static quality report.xlsm").Activate
    Sheets("Metrics").Select
    
    
    ' Writes the variables in the report file, here, the i variable must change from metric to metric (GCA,DRR,DRL,etc.)
    Cells(i + 3, 6).Value = Level2GCA & "-" & Level3GCA & "-" & Level4GCA & "-" & Level5GCA & "-" & FaultGCA
    Cells(i + 3, 11).Value = WDPVGCA
    
   
Next i
'============================================= Screenshot for GCA C1UL CADILAC ==================================================
Windows("GCAParetoNE.xlsx").Activate
' Returns the last row after the cleanse of the data
vCountOfRowsGCA2 = Range("A1", Range("A1").End(xlDown)).Rows.Count

Range(Cells(5, 1), Cells(12, 19)).Select

Application.CutCopyMode = False
Selection.Copy
Selection.CopyPicture xlScreen, xlBitmap
' Pastes the picture in the corresponding range
Windows("SH Static quality report.xlsm").Activate
Worksheets("GCA").Select
Range("D28").Select
Selection.PasteSpecial
' Changes size of image
    With Selection
                    .ShapeRange.LockAspectRatio = msoFalse
                    .ShapeRange.Height = 230
                    .ShapeRange.Width = 1065
    End With

Windows("GCAParetoNE.xlsx").Close SaveChanges:=True

End If
'==============================================================================================================================================
'========================================================== GCA GMC C1UG ======================================================================
'==============================================================================================================================================
third:

FileNameGCAN8 = VBA.Dir("C:\Users\SZ5CM6\Documents\SPHL\GCA\GCAParetoN8.xlsx")

If FileNameGCAN8 = "" Then ' Condition to evaluate if the file exists
MsgBox ("File GCAParetoN8.xlsx does no exists." & vbLf & "Click Ok to continue")
GoTo final

Else
 ' Opens the GCA N8 File in the designated path
Workbooks.Open "C:\Users\SZ5CM6\Documents\SPHL\GCA\" & FileNameGCAN8

Windows("GCAParetoN8.xlsx").Activate

' Returns the last row before the cleanse of the data
vCountOfRowsGCA = Range("A1", Range("A1").End(xlDown)).Rows.Count

' Cycle to only keep "Fits - Exterior"
For a = 6 To vCountOfRowsGCA Step 1
    If Cells(a, 3).Value <> "Fits - Exterior" Then
    
     Rows(a).EntireRow.Clear
    
    End If
Next a

' Orders in descending order, i.e. by WDPV
Range("A5:T10000").Sort Key1:=Range("S6"), Order1:=xlDescending, Header:=xlYes

'============================================= Information transfer for GCA GMC C1UG ==========================================================

' Cycle to initialize the variables for each level and wdpv, and to write each variable in the Report file
For i = 6 To 8 Step 1
 
    Windows("GCAParetoN8.xlsx").Activate
 
    Level2GCA = Cells(i, 5).Value
    Level3GCA = Cells(i, 7).Value
    Level4GCA = Cells(i, 9).Value
    Level5GCA = Cells(i, 11).Value
    FaultGCA = Cells(i, 13).Value
    WDPVGCA = Cells(i, 19).Value
    
    Windows("SH Static quality report.xlsm").Activate
    Sheets("Metrics").Select
    
    
    ' Writes the variables in the report file, here, the i variable must change from metric to metric (GCA,DRR,DRL,etc.)
    Cells(i + 6, 6).Value = Level2GCA & "-" & Level3GCA & "-" & Level4GCA & "-" & Level5GCA & "-" & FaultGCA
    Cells(i + 6, 11).Value = WDPVGCA
    
   
Next i
'============================================= Screenshot for GCA GMC C1UG ==================================================
Windows("GCAParetoN8.xlsx").Activate
' Returns the last row after the cleanse of the data
vCountOfRowsGCA2 = Range("A1", Range("A1").End(xlDown)).Rows.Count

Range(Cells(5, 1), Cells(12, 19)).Select

Application.CutCopyMode = False
Selection.Copy
Selection.CopyPicture xlScreen, xlBitmap
' Pastes the picture in the corresponding range
Windows("SH Static quality report.xlsm").Activate
Worksheets("GCA").Select
Range("D49").Select
Selection.PasteSpecial
' Changes size of image
    With Selection
                    .ShapeRange.LockAspectRatio = msoFalse
                    .ShapeRange.Height = 230
                    .ShapeRange.Width = 1065
    End With

Windows("GCAParetoN8.xlsx").Close SaveChanges:=True
final:
End If
End Sub
