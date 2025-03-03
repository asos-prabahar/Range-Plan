Attribute VB_Name = "Oracle_TextToID_Conversion"
'Cherry change included, Sprint 58, Sprint 68, Version 100, Version 101, Version 102, version 103
Sub Oracle_ID_Conversion()

    'Declaring variables
        Dim ProdMaxRows As Long
        Dim BuyMaxRows As Long
        Dim SupMaxRows As Long
        Dim FacMaxRows As Long
        Dim AttMaxRows As Long
        Dim OrcMaxRows As Long
        Dim OrcMaxCol As Long
        Dim DiffsMaxRows As Long
        Dim RpasDiffsMaxRows As Long
        Dim BrndMaxRows As Long

        Dim OrcColDivision As String
        Dim OrcColDivisionNo As Long
        Dim OrcColSupplier As String
        Dim OrcColFactoryUK As String
        Dim OrcColFactoryEU As String
        Dim OrcColFactoryUS As String
        Dim OrcColColourGrp As String
        Dim OrcColColour As String
        Dim OrcColSizeGrp As String
        Dim OrcColBrand As String
        Dim UDAID_Temp As String

        Dim UDAColNo As Long

'Declared by Prabahar on 02-Sep-2021
    Dim OrcColBusinessModel, OrcColBuyingGrp As String
    Dim OrcColSizeCurveIDUK, OrcColSizeCurveIDEU, OrcColSizeCurveIDUS As String
    Dim OrcColManualSizeDistUK, OrcColManualSizeDistEU, OrcColManualSizeDistUS As String
    Dim i As Byte

    'Inserting temporary columns
        'Columns("A:AH").Select
        Columns("A:AI").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A" & OracleHeadderRow).Select

    'Initializing variables
        ProdMaxRows = Common_Functions.Find_Used_Rows(RPFileName, DATSheetProdHyr)
        BuyMaxRows = Common_Functions.Find_Used_Rows(RPFileName, DATSheetBuyrar)
        SupMaxRows = Common_Functions.Find_Used_Rows(RPFileName, DATSheetSupplier)
        FacMaxRows = Common_Functions.Find_Used_Rows(RPFileName, DATSheetFactory)
        AttMaxRows = Common_Functions.Find_Used_Rows(RPFileName, "AttributesRaw")
        OrcMaxRows = Common_Functions.Find_Used_Rows(RPFileName, OracleSheet)
        OrcMaxCol = Common_Functions.Find_Used_Columns(RPFileName, OracleSheet)
        DiffsMaxRows = Common_Functions.Find_Used_Rows(RPFileName, DATSheetDiffs)
        RpasDiffsMaxRows = Common_Functions.Find_Used_Rows(RPFileName, DATSheetRpasDiffs)
        BrndMaxRows = Common_Functions.Find_Used_Rows(RPFileName, DATSheetBrand)

        OrcColDivision = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "DIVISION")
        OrcColDivisionNo = Find_Specific_Columns_Number(RPFileName, OracleSheet, OracleHeadderRow, "DIVISION")
        OrcColSupplier = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "SUPPLIER SITE")
        OrcColFactoryUK = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "UK FACTORY")
        OrcColFactoryEU = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "EU FACTORY")
        OrcColFactoryUS = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "US FACTORY")
        OrcColColourGrp = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "COLOUR GROUP")
        'Cherry Release
        'OrcColColour = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "COLOUR")
        OrcColColour = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "REPORTING COLOUR")
        OrcColSizeGrp = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "SIZE GROUP")
        OrcColBrand = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "BRAND")

    'Adding Formula
    'If Column Chnages, need to change reference below for 'ProdMaxRows', 'BuyMaxRows','SupMaxRows','FacMaxRows','OrcMaxRows'
        'DIVISION
            Range("A5") = "Division ID"
            Range("A6") = "=VLOOKUP(" & Column_Letter(OrcColDivisionNo) & OracleHeadderRow + 2 & ",CHOOSE({1,2},RpasMerchhier!$J$1:$J$" & ProdMaxRows & ",RpasMerchhier!$I$1:$I$" & ProdMaxRows & "),2,FALSE)"
        'GROUP
            Range("B5") = "Division start"
            Range("B6") = "=MATCH(" & Column_Letter(OrcColDivisionNo) & OracleHeadderRow + 2 & ",RpasMerchhier!$J$1:$J$" & ProdMaxRows & ",0)-1"
            Range("C5") = "Division count"
            Range("C6") = "=COUNTIF(RpasMerchhier!$J$1:$J$" & ProdMaxRows & "," & Column_Letter(OrcColDivisionNo) & OracleHeadderRow + 2 & ")"
            Range("D5") = "Group ID"
            Range("D6") = "=VLOOKUP(" & Column_Letter(OrcColDivisionNo + 1) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$H$1,B6,,C6),OFFSET(RpasMerchhier!$G$1,B6,,C6)),2,FALSE)"
        'PRODUCT GROUP
            Range("E5") = "Group start"
            Range("E6") = "=B6+MATCH(" & Column_Letter(OrcColDivisionNo + 1) & OracleHeadderRow + 2 & ",OFFSET(RpasMerchhier!$H$1,B6,,C6),0)-1"
            Range("F5") = "Group count"
            '13Mar20
            'Range("F6") = "=C6+COUNTIF(OFFSET(RpasMerchhier!$H$1,B6,,C6)," & Column_Letter(OrcColDivisionNo + 1) & OracleHeadderRow + 2 & ")"
            Range("F6") = "=COUNTIF(OFFSET(RpasMerchhier!$H$1,B6,,C6)," & Column_Letter(OrcColDivisionNo + 1) & OracleHeadderRow + 2 & ")"
            Range("G5") = "Product ID"
            Range("G6") = "=VLOOKUP(" & Column_Letter(OrcColDivisionNo + 2) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$F$1,E6,,F6),OFFSET(RpasMerchhier!$E$1,E6,,F6)),2,FALSE)"
        'CATEGORY
            Range("H5") = "Product  Start"
            '13Mar20
            'Range("H6") = "=E6+MATCH(" & Column_Letter(OrcColDivisionNo + 3) & OracleHeadderRow + 2 & ",OFFSET(RpasMerchhier!$D$1,E6,,F6),0)-1"
            Range("H6") = "=E6+MATCH(" & Column_Letter(OrcColDivisionNo + 2) & OracleHeadderRow + 2 & ",OFFSET(RpasMerchhier!$F$1,E6,,F6),0)-1"
            Range("I5") = "Product   Count"
            '13Mar20
            'Range("I6") = "=F6+COUNTIF(OFFSET(RpasMerchhier!$D$1,E6,,F6)," & Column_Letter(OrcColDivisionNo + 3) & OracleHeadderRow + 2 & ")"
            Range("I6") = "=COUNTIF(OFFSET(RpasMerchhier!$F$1,E6,,F6)," & Column_Letter(OrcColDivisionNo + 2) & OracleHeadderRow + 2 & ")"
            'Range("J6") = "=VLOOKUP(" & Column_Letter(OrcColDivisionNo + 3) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$D$1,H6,,I6),OFFSET(RpasMerchhier!$C$1,H6,,I6)),2,FALSE)"'Sprint 68
            Range("J5") = "Category ID"
            Range("J6") = "=IFERROR(LEFT(VLOOKUP(" & Column_Letter(OrcColDivisionNo + 3) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$D$1,H6,,I6),OFFSET(RpasMerchhier!$C$1,H6,,I6)),2,FALSE),FIND(""_"",VLOOKUP(" & Column_Letter(OrcColDivisionNo + 3) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$D$1,H6,,I6),OFFSET(RpasMerchhier!$C$1,H6,,I6)),2,FALSE),1)-1),"""")"
        'SUB-CATEGORY
            Range("K5") = "Sub Category Start"
            '13Mar20
            'Range("K6") = "=H6+MATCH(" & Column_Letter(OrcColDivisionNo + 4) & OracleHeadderRow + 2 & ",OFFSET(RpasMerchhier!$B$1,H6,,I6),0)-1"
            Range("K6") = "=H6+MATCH(" & Column_Letter(OrcColDivisionNo + 3) & OracleHeadderRow + 2 & ",OFFSET(RpasMerchhier!$D$1,H6,,I6),0)-1"
            Range("L5") = "Sub Category Count"
            '13Mar20
            'Range("L6") = "=I6+COUNTIF(OFFSET(RpasMerchhier!$B$1,H6,,I6)," & Column_Letter(OrcColDivisionNo + 4) & OracleHeadderRow + 2 & ")"
            Range("L6") = "=COUNTIF(OFFSET(RpasMerchhier!$D$1,H6,,I6)," & Column_Letter(OrcColDivisionNo + 3) & OracleHeadderRow + 2 & ")"
            Range("M5") = "Sub Cat ID"
            'Range("M6") = "=VLOOKUP(" & Column_Letter(OrcColDivisionNo + 4) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$B$1,K6,,L6),OFFSET(RpasMerchhier!$A$1,K6,,L6)),2,FALSE)" 'Sprint 68
            Range("M6") = "=IFERROR(LEFT(VLOOKUP(" & Column_Letter(OrcColDivisionNo + 4) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$B$1,K6,,L6),OFFSET(RpasMerchhier!$A$1,K6,,L6)),2,FALSE),FIND(""_"",VLOOKUP(" & Column_Letter(OrcColDivisionNo + 4) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$B$1,K6,,L6),OFFSET(RpasMerchhier!$A$1,K6,,L6)),2,FALSE),1)-1),"""")"
        'BUSINESS MODEL
            Range("N5") = "Business Model ID"
            Range("N6") = "=VLOOKUP(" & Column_Letter(OrcColDivisionNo + 5) & OracleHeadderRow + 2 & ",CHOOSE({1,2},Buyrachy!$B$1:$B$" & BuyMaxRows & ",Buyrachy!$A$1:$A$" & BuyMaxRows & "),2,FALSE)"
        'BUYING GROUP
            Range("O5") = "Business Model Start"
            Range("O6") = "=MATCH(" & Column_Letter(OrcColDivisionNo + 5) & OracleHeadderRow + 2 & ",Buyrachy!$B$1:$B$" & BuyMaxRows & ",0)-1"
            Range("P5") = "Business Model Count"
            Range("P6") = "=COUNTIF(Buyrachy!$B$1:$B$" & BuyMaxRows & "," & Column_Letter(OrcColDivisionNo + 5) & OracleHeadderRow + 2 & ")"
            Range("Q5") = "Buying Group ID"
            'Range("Q6") = "=VLOOKUP(" & Column_Letter(OrcColDivisionNo + 6) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$D$1,O6,,P6),OFFSET(Buyrachy!$C$1,O6,,P6)),2,FALSE)"'Sprint 68
            Range("Q6") = "=IFERROR(LEFT(VLOOKUP(" & Column_Letter(OrcColDivisionNo + 6) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$D$1,O6,,P6),OFFSET(Buyrachy!$C$1,O6,,P6)),2,FALSE),FIND(""_"",VLOOKUP(" & Column_Letter(OrcColDivisionNo + 6) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$D$1,O6,,P6),OFFSET(Buyrachy!$C$1,O6,,P6)),2,FALSE),1)-1),"""")"
        'BUYING SUB GROUP
            Range("R5") = "Buying Group Start"
            Range("R6") = "=O6+MATCH(" & Column_Letter(OrcColDivisionNo + 6) & OracleHeadderRow + 2 & ",OFFSET(Buyrachy!$D$1,O6,,P6),0)-1"
            '13Mar20
            Range("S5") = "Buying Group Count"
            'Range("S6") = "=P6+COUNTIF(OFFSET(Buyrachy!$D$1,O6,,P6)," & Column_Letter(OrcColDivisionNo + 6) & OracleHeadderRow + 2 & ")"
            Range("S6") = "=COUNTIF(OFFSET(Buyrachy!$D$1,O6,,P6)," & Column_Letter(OrcColDivisionNo + 6) & OracleHeadderRow + 2 & ")"
            Range("T5") = "Buying SubGroup ID"
            'Range("T6") = "=VLOOKUP(" & Column_Letter(OrcColDivisionNo + 7) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$F$1,R6,,S6),OFFSET(Buyrachy!$E$1,R6,,S6)),2,FALSE)"'Sprint 68
            Range("T6") = "=IFERROR(LEFT(VLOOKUP(" & Column_Letter(OrcColDivisionNo + 7) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$F$1,R6,,S6),OFFSET(Buyrachy!$E$1,R6,,S6)),2,FALSE),FIND(""_"",VLOOKUP(" & Column_Letter(OrcColDivisionNo + 7) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$F$1,R6,,S6),OFFSET(Buyrachy!$E$1,R6,,S6)),2,FALSE),1)-1),"""")"
        'BUYING SET
            Range("U5") = "Buying SubGroup Start"
            Range("U6") = "=R6+MATCH(" & Column_Letter(OrcColDivisionNo + 7) & OracleHeadderRow + 2 & ",OFFSET(Buyrachy!$F$1,R6,,S6),0)-1"
            '13Mar20
            Range("V5") = "Buying SubGroup Count"
            'Range("V6") = "=S6+COUNTIF(OFFSET(Buyrachy!$F$1,R6,,S6)," & Column_Letter(OrcColDivisionNo + 7) & OracleHeadderRow + 2 & ")"
            Range("V6") = "=COUNTIF(OFFSET(Buyrachy!$F$1,R6,,S6)," & Column_Letter(OrcColDivisionNo + 7) & OracleHeadderRow + 2 & ")"
            Range("W5") = "Buying Set ID"
            'Range("W6") = "=VLOOKUP(" & Column_Letter(OrcColDivisionNo + 8) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$H$1,U6,,V6),OFFSET(Buyrachy!$G$1,U6,,V6)),2,FALSE)"'Sprint 68
            Range("W6") = "=IFERROR(LEFT(VLOOKUP(" & Column_Letter(OrcColDivisionNo + 8) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$H$1,U6,,V6),OFFSET(Buyrachy!$G$1,U6,,V6)),2,FALSE),FIND(""_"",VLOOKUP(" & Column_Letter(OrcColDivisionNo + 8) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$H$1,U6,,V6),OFFSET(Buyrachy!$G$1,U6,,V6)),2,FALSE),1)-1),"""")"
        'SUPPLIER SITE
            Range("X5") = "Supplier ID"
            Range("X6") = "=VLOOKUP(" & OrcColSupplier & OracleHeadderRow + 2 & ",CHOOSE({1,2},RpasSuppliers!$B$1:$B$" & SupMaxRows & ",RpasSuppliers!$A$1:$A$" & SupMaxRows & "),2,FALSE)"
        'FACTORY
            Range("Y5") = "Supplier ID Start"
            Range("Y6") = "=MATCH(X6,SuppliersFactories!$A$1:$A$" & FacMaxRows & ",0)-1"
            Range("Z5") = "Supplier ID Count"
            Range("Z6") = "=COUNTIF(SuppliersFactories!$A$1:$A$" & FacMaxRows & ",X6)"
            Range("AA5") = "Factory ID"
            Range("AA6") = "=VLOOKUP(" & OrcColFactoryUK & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(SuppliersFactories!$C$1,Y6,,Z6),OFFSET(SuppliersFactories!$B$1,Y6,,Z6)),2,FALSE)"
        'COLOUR GROUP
            Range("AB5") = "Colour Group ID"
            Range("AB6") = "=VLOOKUP(" & OrcColColourGrp & OracleHeadderRow + 2 & ",CHOOSE({1,2},Diffs!$B$1:$B$" & DiffsMaxRows & ",Diffs!$A$1:$A$" & DiffsMaxRows & "),2,FALSE)"
        'COLOUR
            Range("AC5") = "Colour Group ID Start"
            Range("AC6") = "=MATCH(AB6,RpasDiffs!$E$1:$E$" & RpasDiffsMaxRows & ",0)-1"
            Range("AD5") = "Colour Group ID Count"
            Range("AD6") = "=COUNTIF(RpasDiffs!$E$1:$E$" & RpasDiffsMaxRows & ",AB6)"
            Range("AE5") = "Colour (Oracle) ID"
            Range("AE6") = "=VLOOKUP(" & OrcColColour & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasDiffs!$B$1,AC6,,AD6),OFFSET(RpasDiffs!$A$1,AC6,,AD6)),2,FALSE)"
        'SIZE  GROUP
            Range("AF5") = "Size Data Start"
            Range("AF6") = "=MATCH(""Size"",Diffs!$D$1:$D$" & DiffsMaxRows & ",0)-1"
            Range("AG5") = "Size Data Count"
            Range("AG6") = "=COUNTIF(Diffs!$D$1:$D$" & DiffsMaxRows & ",""Size"")"

        'Cherry release
        'Range("AH6") = "=VLOOKUP(" & OrcColSizeGrp & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasDiffs!$B$1,AF6,,AG6),OFFSET(RpasDiffs!$A$1,AF6,,AG6)),2,FALSE)"
            Range("AH5") = "Size Group ID"
            Range("AH6") = "=VLOOKUP(" & OrcColSizeGrp & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Diffs!$H$1,AF6,,AG6),OFFSET(Diffs!$G$1,AF6,,AG6)),2,FALSE)"

        'FACTORY
            Range("AI5") = "Brand ID"
            Range("AI6") = "=VLOOKUP(" & OrcColBrand & OracleHeadderRow + 2 & ",CHOOSE({1,2},Brands!$B$1:$B$" & BrndMaxRows & ",Brands!$A$1:$A$" & BrndMaxRows & "),2,FALSE)"

    'Copypasting and Dragging Formula
        Range("A" & OracleHeadderRow + 2 & ":AI" & OracleHeadderRow + 2).Copy
        Range("A" & OracleHeadderRow + 2 & ":AI" & OrcMaxRows).Select
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        Selection.Replace What:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        Range("A" & OracleHeadderRow).Select

    'Pasting Product Hyrarchy columns IDs and Buyrarchy columns IDs into respective columns
        'Range("A" & OracleHeadderRow + 2 & ":A" & OrcMaxRows & ",D" & OracleHeadderRow + 2 & ":D" & OrcMaxRows & ",G" & OracleHeadderRow + 2 & ":G" & OrcMaxRows & ",J" & OracleHeadderRow + 2 & ":J" & OrcMaxRows & ",M" & OracleHeadderRow + 2 & ":M" & OrcMaxRows & ",N" & OracleHeadderRow + 2 & ":N" & OrcMaxRows & ",Q" & OracleHeadderRow + 2 & ":Q" & OrcMaxRows & ",T" & OracleHeadderRow + 2 & ":T" & OrcMaxRows & ",W" & OracleHeadderRow + 2 & ":W" & OrcMaxRows).Copy

        'Cherry release --- Changes Start
        'Range("A" & OracleHeadderRow + 2 & ":A" & OrcMaxRows & ",D" & OracleHeadderRow + 2 & ":D" & OrcMaxRows & ",G" & OracleHeadderRow + 2 & ":G" & OrcMaxRows & ",J" & OracleHeadderRow + 2 & ":J" & OrcMaxRows & ",M" & OracleHeadderRow + 2 & ":M" & OrcMaxRows & ",N" & OracleHeadderRow + 2 & ":N" & OrcMaxRows & ",Q" & OracleHeadderRow + 2 & ":Q" & OrcMaxRows & ",T" & OracleHeadderRow + 2 & ":T" & OrcMaxRows & ",W" & OracleHeadderRow + 2 & ":W" & OrcMaxRows & ",AI" & OracleHeadderRow + 2 & ":AI" & OrcMaxRows).Copy
        'Range(OrcColDivision & OracleHeadderRow + 2).Select
        'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        Range("A" & OracleHeadderRow + 2 & ":A" & OrcMaxRows & ",D" & OracleHeadderRow + 2 & ":D" & OrcMaxRows & ",G" & OracleHeadderRow + 2 & ":G" & OrcMaxRows & ",J" & OracleHeadderRow + 2 & ":J" & OrcMaxRows & ",M" & OracleHeadderRow + 2 & ":M" & OrcMaxRows & ",N" & OracleHeadderRow + 2 & ":N" & OrcMaxRows & ",Q" & OracleHeadderRow + 2 & ":Q" & OrcMaxRows & ",T" & OracleHeadderRow + 2 & ":T" & OrcMaxRows & ",W" & OracleHeadderRow + 2 & ":W" & OrcMaxRows).Copy
        Range(OrcColDivision & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        Range("AI" & OracleHeadderRow + 2 & ":AI" & OrcMaxRows).Copy
        Range(OrcColBrand & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        'Cherry release --- Changes End

    'Pasting Supplier IDs
        Range("X" & OracleHeadderRow + 2 & ":X" & OrcMaxRows).Copy
        Range(OrcColSupplier & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    'Pasting 3 Factory Columns
        Range("AA" & OracleHeadderRow + 2 & ":AA" & OrcMaxRows).Copy
        Range(OrcColFactoryUK & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Range(OrcColFactoryEU & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Range(OrcColFactoryUS & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    'Pasting Colour Group, Colour And Size Group Columns
        Range("AB" & OracleHeadderRow + 2 & ":AB" & OrcMaxRows).Copy
        Range(OrcColColourGrp & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        Range("AE" & OracleHeadderRow + 2 & ":AE" & OrcMaxRows).Copy
        Range(OrcColColour & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        Range("AH" & OracleHeadderRow + 2 & ":AH" & OrcMaxRows).Copy
        Range(OrcColSizeGrp & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    '==== Processing UDA Columns

    'Deleting few temporary columns
        'Columns("A:AD").Select
        Columns("A:AE").Select
        Selection.Delete Shift:=xlToLeft
        Range("A" & OracleHeadderRow).Select

        UDAColNo = 1

        While UDAColNo <= OrcMaxCol
            If Left(Range(Column_Letter(UDAColNo) & OracleHeadderRow).Text, 3) = "UDA" Then

                    'Range("A6") = "=VLOOKUP(" & Column_Letter(UDAColNo) & OracleHeadderRow + 2 & ",CHOOSE({1,2},AttributesRaw!$B$1:$B$" & AttMaxRows & ",AttributesRaw!$A$1:$A$" & AttMaxRows & "),2,FALSE)"
                    'Range("B6") = "=MATCH(" & Column_Letter(UDAColNo) & OracleHeadderRow + 2 & ",AttributesRaw!$B$1:$B$" & AttMaxRows & ",0)-1"
                    'Range("C6") = "=COUNTIF(AttributesRaw!$B$1:$B$" & AttMaxRows & "," & Column_Letter(UDAColNo) & OracleHeadderRow + 2 & ")"

                    'Sprint 58
                    'Range("A6") = Mid(Range(Column_Letter(UDAColNo) & OracleHeadderRow), 8, 10)
                    UDAID_Temp = Mid(Range(Column_Letter(UDAColNo) & OracleHeadderRow), 8, 10) 'Updated on 4-Nov-2020
                    Range("A6") = "=IF(" & Column_Letter(UDAColNo + 1) & OracleHeadderRow + 2 & "= " & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & UDAID_Temp & ")" 'Updated on 4-Nov-2020

                    Range("B6") = "=MATCH(" & UDAID_Temp & ",AttributesRaw!$A$1:$A$" & AttMaxRows & ",0)-1" 'updated by Salih 30-Mar-2021
                    Range("C6") = "=COUNTIF(AttributesRaw!$A$1:$A$" & AttMaxRows & "," & UDAID_Temp & ")"   'updated by Salih 30-Mar-2021
                    Range("B6") = Range("B6").Text
                    Range("C6") = Range("C6").Text

                    'Checking if the field is dropdown to convert into IDs. If not dropdown then conversion will not happen
                    Range("D6") = "=VLOOKUP(" & UDAID_Temp & ",AttributesRaw!A:C,3,0)" 'updated by Salih 30-Mar-2021
                    If Range("D6").Text = "LV" Then
                        'Range("D6") = "=VLOOKUP(" & Column_Letter(UDAColNo + 1) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(AttributesRaw!$H$1,B6,,C6),OFFSET(AttributesRaw!$G$1,B6,,C6)),2,FALSE)"
                        Range("D6") = "=VLOOKUP(" & Column_Letter(UDAColNo + 1) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(AttributesRaw!$H$1,B6,,C6),OFFSET(AttributesRaw!$G$1,B6,,C6)),2,FALSE)"
                    ElseIf Range("D6").Text = "DT" Then
                        Range("D6") = "=TEXT(" & Column_Letter(UDAColNo + 1) & OracleHeadderRow + 2 & ",""dd/mm/yyyy"")"
                    Else
                        Range("D6") = "=" & Column_Letter(UDAColNo + 1) & OracleHeadderRow + 2
                    End If

                'Copypasting and Dragging Formula
                    Range("A" & OracleHeadderRow + 2 & ":D" & OracleHeadderRow + 2).Copy
                    Range("A" & OracleHeadderRow + 2 & ":D" & OrcMaxRows).Select
                    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                    Selection.Copy
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                    Range("A" & OracleHeadderRow).Select

                'Copypasting IDs into respective column
                    Range("A" & OracleHeadderRow + 2 & ":A" & OrcMaxRows & ",D" & OracleHeadderRow + 2 & ":D" & OrcMaxRows).Copy
                    Range(Column_Letter(UDAColNo) & OracleHeadderRow + 2).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                    Selection.Replace What:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                    Range("A" & OracleHeadderRow).Select
                    Application.CutCopyMode = False

                    UDAColNo = UDAColNo + 1
                End If
            UDAColNo = UDAColNo + 1
        Wend

'This block converts the size curves into IDs  | Prabahar 02-Sep-2021


'Finding specific columns for the respective header names
    OrcColBusinessModel = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "BUSINESS MODEL")
    OrcColBuyingGrp = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "BUYING GROUP")
    OrcColSizeGrp = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "SIZE GROUP")
    OrcColSizeCurveIDUK = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "UK SIZE CURVE ID")
    OrcColSizeCurveIDEU = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "EU SIZE CURVE ID")
    OrcColSizeCurveIDUS = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "US SIZE CURVE ID")
    OrcColManualSizeDistUK = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "UK MANUAL SIZE DIST")
    OrcColManualSizeDistEU = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "EU MANUAL SIZE DIST")
    OrcColManualSizeDistUS = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "US MANUAL SIZE DIST")

'Copypasting and Dragging Formula
    Range("A" & OracleHeadderRow + 2) = "=IFERROR(VLOOKUP(" & OrcColSizeCurveIDUK & OracleHeadderRow + 2 & "&""_""&" & OrcColSizeGrp & OracleHeadderRow + 2 & "&""_""&" & OrcColBuyingGrp & OracleHeadderRow + 2 & "&""_""&" & OrcColBusinessModel & OracleHeadderRow + 2 & ",CascadingDropdowns!BH:BI,2,0),"""")"
    Range("B" & OracleHeadderRow + 2) = "=IFERROR(VLOOKUP(" & OrcColSizeCurveIDEU & OracleHeadderRow + 2 & "&""_""&" & OrcColSizeGrp & OracleHeadderRow + 2 & "&""_""&" & OrcColBuyingGrp & OracleHeadderRow + 2 & "&""_""&" & OrcColBusinessModel & OracleHeadderRow + 2 & ",CascadingDropdowns!BH:BI,2,0),"""")"
    Range("C" & OracleHeadderRow + 2) = "=IFERROR(VLOOKUP(" & OrcColSizeCurveIDUS & OracleHeadderRow + 2 & "&""_""&" & OrcColSizeGrp & OracleHeadderRow + 2 & "&""_""&" & OrcColBuyingGrp & OracleHeadderRow + 2 & "&""_""&" & OrcColBusinessModel & OracleHeadderRow + 2 & ",CascadingDropdowns!BH:BI,2,0),"""")"

    Range("A" & OracleHeadderRow + 2 & ":C" & OracleHeadderRow + 2).Copy
    Range("A" & OracleHeadderRow + 2 & ":C" & OrcMaxRows).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

'Moving Size Curve IDs to respective columns
    For i = 1 To 3
        If i = 1 Then
            Range(OrcColSizeCurveIDUK & OracleHeadderRow + 2) = "=IF(" & OrcColManualSizeDistUK & OracleHeadderRow + 2 & " = """", IF(A6="""","""",1" & "&A6),"""")"
            Range(OrcColSizeCurveIDUK & OracleHeadderRow + 2).Copy
            Range(OrcColSizeCurveIDUK & OracleHeadderRow + 2 & ":" & OrcColSizeCurveIDUK & OrcMaxRows).Select
        ElseIf i = 2 Then
            Range(OrcColSizeCurveIDEU & OracleHeadderRow + 2) = "=IF(" & OrcColManualSizeDistEU & OracleHeadderRow + 2 & " = """", IF(B6="""","""",4" & "&B6),"""")"
            Range(OrcColSizeCurveIDEU & OracleHeadderRow + 2).Copy
            Range(OrcColSizeCurveIDEU & OracleHeadderRow + 2 & ":" & OrcColSizeCurveIDEU & OrcMaxRows).Select
        ElseIf i = 3 Then
            Range(OrcColSizeCurveIDUS & OracleHeadderRow + 2) = "=IF(" & OrcColManualSizeDistUS & OracleHeadderRow + 2 & " = """", IF(C6="""","""",3" & "&C6),"""")"
            Range(OrcColSizeCurveIDUS & OracleHeadderRow + 2).Copy
            Range(OrcColSizeCurveIDUS & OracleHeadderRow + 2 & ":" & OrcColSizeCurveIDUS & OrcMaxRows).Select
        End If
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    Next i

    Range("A" & OracleHeadderRow).Select

    'Deleting remaining temporary columns
        Columns("A:D").Select
        Selection.Delete Shift:=xlToLeft
        Range("A" & OracleHeadderRow).Select
        Selection.Find(What:="*", LookAt:=xlPart).Activate
End Sub

Function MacroVersionCheck_Oracle_TextToID_Conversion() As Boolean
    If Sheets(ManRefDataSheet).Range("AB18") = "Oracle_TextToID_Conversion_v103" Then MacroVersionCheck_Oracle_TextToID_Conversion = True
End Function

