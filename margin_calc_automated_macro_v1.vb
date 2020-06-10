Option Explicit

Sub margin_calculation_automated_v01()

   Dim maxRow1 As Long
   Dim maxRow2 As Long
   Dim maxRow1Address As String
   Dim maxRow2Address As String
   Dim lRow As Long
   Dim lCol As Integer
   
   Dim intCol As Integer
   Dim lngRow As Long
   
   Dim wartMargin As Single
   Dim procMargin As Single
   Dim country As String
   
   ActiveSheet.Range("A" & Rows.Count).End(xlUp).Select       'line finds last used cell in column "A"
   Debug.Print ActiveSheet.Range("A" & Rows.Count).End(xlUp).Address
   lRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).row
   
   ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Select   'line finds last used cell in 1st row
   Debug.Print ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Address
   lCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
   
   maxRow1 = ActiveSheet.Range("A" & Rows.Count).End(xlUp)    'line returns referenced range .Value, not the address.
   maxRow1Address = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Address
   Debug.Print maxRow1Address
   
   maxRow2 = ActiveSheet.Range("F" & Rows.Count).End(xlUp)    'line returns referenced range .Value, not the address.
   maxRow2Address = ActiveSheet.Range("E" & Rows.Count).End(xlUp).Address
   Debug.Print maxRow2Address

   MsgBox "Last Row: " & lRow & vbNewLine & _
          "Last Column: " & Columns(lCol).Address(False, False)
   
   ActiveSheet.Range("H2:H" & lRow).Formula = "=F2-G2"
   
   ActiveSheet.Range("I2:I" & lRow).Formula = "=H2/F2"
   
   ActiveSheet.Range("J2:J" & lRow).FormulaR1C1 = "=VLOOKUP(RC4,'Tabela Kraj'!R1C1:R679C2,2,0)"
   
   
   MsgBox "Teraz zrobimy sortowanie po kolumnie z marżą procentową."
   
   intCol = Application.WorksheetFunction.Match("%Marży", Worksheets("Dane Zadanie 2").Range("1:1"), 0)    'Range("1:1") is row 1.
   Range("A1").CurrentRegion.Sort Key1:=ActiveSheet.Cells(1, intCol), Order1:=xlAscending, Header:=xlYes
   
   
   'the code lines below come from "update_custdb_v2" macro;
   'lngRow = Address.row
   'intCol = Application.WorksheetFunction.Match("data_waznosci_slownie", Worksheets("data").Range("1:1"), 0)    'Range("1:1") is row 1.
   'mySheet.Cells(lngRow, intCol) = nonFormatedDate
   
End Sub
