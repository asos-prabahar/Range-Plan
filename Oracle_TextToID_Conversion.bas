Option Explicit
'Cherry change included, Sprint 58, Sprint 68, Version 100, Version 101, Version 102, version 103,104,105,106,107
Sub Oracle_ID_Conversion()

    'Declaring variables
        Dim prodMaxRows As Long
        Dim buyMaxRows As Long
        Dim supMaxRows As Long
        Dim facMaxRows As Long
        Dim attMaxRows As Long
        Dim orcMaxRows As Long
        Dim orcMaxCol As Long
        Dim diffsMaxRows As Long
        Dim rpasDiffsMaxRows As Long
        Dim brndMaxRows As Long
        Dim poPriorityRow As Long
        Dim priorityFreq As Long

        Dim orcColDivision As String
        Dim orcColDivisionNo As Long
        Dim orcColSupplier As String
        Dim orcColFactoryUK As String
        Dim orcColFactoryEU As String
        Dim orcColFactoryUS As String
        Dim orcColColourGrp As String
        Dim orcColColour As String
        Dim orcColSizeGrp As String
        Dim orcColBrand As String
        Dim UDAID_Temp As String

        Dim UDAColNo As Long

    'Declared by Prabahar on 02-Sep-2021
        Dim orcColBusinessModel, orcColBuyingGrp As String
        Dim orcColSizeCurveIDUK, orcColSizeCurveIDEU, orcColSizeCurveIDUS As String
        Dim orcColManualSizeDistUK, orcColManualSizeDistEU, orcColManualSizeDistUS As String
        Dim orcColUKPoPriority, orcColEUPoPriority, orcColUSPoPriority As String 'Added by Manu on 30-Oct-2023
        Dim i As Byte

    'Variables used to for Text-To-ID converstion process for the shipping combination  'Prabahar 27-Fab-2025
        Dim orcUKSPCol As String
        Dim orcUKSMCol As String
        Dim orcUKFFCol As String
        Dim orcEUSPCol As String
        Dim orcEUSMCol As String
        Dim orcEUFFCol As String
        Dim orcUSSPCol As String
        Dim orcUSSMCol As String
        Dim orcUSFFCol As String
        Dim DATSheetFreightMat As String

    Call Initialize_Global_Variables

    'Inserting temporary columns
        'Columns("A:AH").Select
        Columns("A:AI").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A" & OracleHeadderRow).Select

    'Initializing variables
        DATSheetFreightMat = "FreightMatrixRaw"

        prodMaxRows = Common_Functions.Find_Used_Rows(RPFileName, DATSheetProdHyr)
        buyMaxRows = Common_Functions.Find_Used_Rows(RPFileName, DATSheetBuyrar)
        supMaxRows = Common_Functions.Find_Used_Rows(RPFileName, DATSheetSupplier)
        facMaxRows = Common_Functions.Find_Used_Rows(RPFileName, DATSheetFactory)
        attMaxRows = Common_Functions.Find_Used_Rows(RPFileName, "AttributesRaw")
        orcMaxRows = Common_Functions.Find_Used_Rows(RPFileName, OracleSheet)
        orcMaxCol = Common_Functions.Find_Used_Columns(RPFileName, OracleSheet)
        diffsMaxRows = Common_Functions.Find_Used_Rows(RPFileName, DATSheetDiffs)
        rpasDiffsMaxRows = Common_Functions.Find_Used_Rows(RPFileName, DATSheetRpasDiffs)
        brndMaxRows = Common_Functions.Find_Used_Rows(RPFileName, DATSheetBrand)


        orcColDivision = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "DIVISION")
        orcColDivisionNo = Find_Specific_Columns_Number(RPFileName, OracleSheet, OracleHeadderRow, "DIVISION")
        orcColSupplier = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "SUPPLIER SITE")
        orcColFactoryUK = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "UK FACTORY")
        orcColFactoryEU = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "EU FACTORY")
        orcColFactoryUS = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "US FACTORY")
        orcColColourGrp = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "COLOUR GROUP")
        'Cherry Release
        'OrcColColour = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "COLOUR")
        orcColColour = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "REPORTING COLOUR")
        orcColSizeGrp = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "SIZE GROUP")
        orcColBrand = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "BRAND")
        orcColUKPoPriority = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "UK PO PRIORITY") 'Added by Manu on 30-Oct-2023
        orcColEUPoPriority = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "EU PO PRIORITY") 'Added by Manu on 30-Oct-2023
        orcColUSPoPriority = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "US PO PRIORITY") 'Added by Manu on 30-Oct-2023

    'Adding Formula
    'If Column Chnages, need to change reference below for 'ProdMaxRows', 'BuyMaxRows','SupMaxRows','FacMaxRows','OrcMaxRows'
        'DIVISION
            Range("A5") = "Division ID"
            Range("A6") = "=VLOOKUP(" & Column_Letter(orcColDivisionNo) & OracleHeadderRow + 2 & ",CHOOSE({1,2},RpasMerchhier!$J$1:$J$" & prodMaxRows & ",RpasMerchhier!$I$1:$I$" & prodMaxRows & "),2,FALSE)"
        'GROUP
            Range("B5") = "Division start"
            Range("B6") = "=MATCH(" & Column_Letter(orcColDivisionNo) & OracleHeadderRow + 2 & ",RpasMerchhier!$J$1:$J$" & prodMaxRows & ",0)-1"
            Range("C5") = "Division count"
            Range("C6") = "=COUNTIF(RpasMerchhier!$J$1:$J$" & prodMaxRows & "," & Column_Letter(orcColDivisionNo) & OracleHeadderRow + 2 & ")"
            Range("D5") = "Group ID"
            Range("D6") = "=VLOOKUP(" & Column_Letter(orcColDivisionNo + 1) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$H$1,B6,,C6),OFFSET(RpasMerchhier!$G$1,B6,,C6)),2,FALSE)"
        'PRODUCT GROUP
            Range("E5") = "Group start"
            Range("E6") = "=B6+MATCH(" & Column_Letter(orcColDivisionNo + 1) & OracleHeadderRow + 2 & ",OFFSET(RpasMerchhier!$H$1,B6,,C6),0)-1"
            Range("F5") = "Group count"
            '13Mar20
            'Range("F6") = "=C6+COUNTIF(OFFSET(RpasMerchhier!$H$1,B6,,C6)," & Column_Letter(OrcColDivisionNo + 1) & OracleHeadderRow + 2 & ")"
            Range("F6") = "=COUNTIF(OFFSET(RpasMerchhier!$H$1,B6,,C6)," & Column_Letter(orcColDivisionNo + 1) & OracleHeadderRow + 2 & ")"
            Range("G5") = "Product ID"
            Range("G6") = "=VLOOKUP(" & Column_Letter(orcColDivisionNo + 2) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$F$1,E6,,F6),OFFSET(RpasMerchhier!$E$1,E6,,F6)),2,FALSE)"
        'CATEGORY
            Range("H5") = "Product  Start"
            '13Mar20
            'Range("H6") = "=E6+MATCH(" & Column_Letter(OrcColDivisionNo + 3) & OracleHeadderRow + 2 & ",OFFSET(RpasMerchhier!$D$1,E6,,F6),0)-1"
            Range("H6") = "=E6+MATCH(" & Column_Letter(orcColDivisionNo + 2) & OracleHeadderRow + 2 & ",OFFSET(RpasMerchhier!$F$1,E6,,F6),0)-1"
            Range("I5") = "Product   Count"
            '13Mar20
            'Range("I6") = "=F6+COUNTIF(OFFSET(RpasMerchhier!$D$1,E6,,F6)," & Column_Letter(OrcColDivisionNo + 3) & OracleHeadderRow + 2 & ")"
            Range("I6") = "=COUNTIF(OFFSET(RpasMerchhier!$F$1,E6,,F6)," & Column_Letter(orcColDivisionNo + 2) & OracleHeadderRow + 2 & ")"
            'Range("J6") = "=VLOOKUP(" & Column_Letter(OrcColDivisionNo + 3) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$D$1,H6,,I6),OFFSET(RpasMerchhier!$C$1,H6,,I6)),2,FALSE)"'Sprint 68
            Range("J5") = "Category ID"
            Range("J6") = "=IFERROR(LEFT(VLOOKUP(" & Column_Letter(orcColDivisionNo + 3) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$D$1,H6,,I6),OFFSET(RpasMerchhier!$C$1,H6,,I6)),2,FALSE),FIND(""_"",VLOOKUP(" & Column_Letter(orcColDivisionNo + 3) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$D$1,H6,,I6),OFFSET(RpasMerchhier!$C$1,H6,,I6)),2,FALSE),1)-1),"""")"
        'SUB-CATEGORY
            Range("K5") = "Sub Category Start"
            '13Mar20
            'Range("K6") = "=H6+MATCH(" & Column_Letter(OrcColDivisionNo + 4) & OracleHeadderRow + 2 & ",OFFSET(RpasMerchhier!$B$1,H6,,I6),0)-1"
            Range("K6") = "=H6+MATCH(" & Column_Letter(orcColDivisionNo + 3) & OracleHeadderRow + 2 & ",OFFSET(RpasMerchhier!$D$1,H6,,I6),0)-1"
            Range("L5") = "Sub Category Count"
            '13Mar20
            'Range("L6") = "=I6+COUNTIF(OFFSET(RpasMerchhier!$B$1,H6,,I6)," & Column_Letter(OrcColDivisionNo + 4) & OracleHeadderRow + 2 & ")"
            Range("L6") = "=COUNTIF(OFFSET(RpasMerchhier!$D$1,H6,,I6)," & Column_Letter(orcColDivisionNo + 3) & OracleHeadderRow + 2 & ")"
            Range("M5") = "Sub Cat ID"
            'Range("M6") = "=VLOOKUP(" & Column_Letter(OrcColDivisionNo + 4) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$B$1,K6,,L6),OFFSET(RpasMerchhier!$A$1,K6,,L6)),2,FALSE)" 'Sprint 68
            Range("M6") = "=IFERROR(LEFT(VLOOKUP(" & Column_Letter(orcColDivisionNo + 4) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$B$1,K6,,L6),OFFSET(RpasMerchhier!$A$1,K6,,L6)),2,FALSE),FIND(""_"",VLOOKUP(" & Column_Letter(orcColDivisionNo + 4) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasMerchhier!$B$1,K6,,L6),OFFSET(RpasMerchhier!$A$1,K6,,L6)),2,FALSE),1)-1),"""")"
        'BUSINESS MODEL
            Range("N5") = "Business Model ID"
            Range("N6") = "=VLOOKUP(" & Column_Letter(orcColDivisionNo + 5) & OracleHeadderRow + 2 & ",CHOOSE({1,2},Buyrachy!$B$1:$B$" & buyMaxRows & ",Buyrachy!$A$1:$A$" & buyMaxRows & "),2,FALSE)"
        'BUYING GROUP
            Range("O5") = "Business Model Start"
            Range("O6") = "=MATCH(" & Column_Letter(orcColDivisionNo + 5) & OracleHeadderRow + 2 & ",Buyrachy!$B$1:$B$" & buyMaxRows & ",0)-1"
            Range("P5") = "Business Model Count"
            Range("P6") = "=COUNTIF(Buyrachy!$B$1:$B$" & buyMaxRows & "," & Column_Letter(orcColDivisionNo + 5) & OracleHeadderRow + 2 & ")"
            Range("Q5") = "Buying Group ID"
            'Range("Q6") = "=VLOOKUP(" & Column_Letter(OrcColDivisionNo + 6) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$D$1,O6,,P6),OFFSET(Buyrachy!$C$1,O6,,P6)),2,FALSE)"'Sprint 68
            Range("Q6") = "=IFERROR(LEFT(VLOOKUP(" & Column_Letter(orcColDivisionNo + 6) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$D$1,O6,,P6),OFFSET(Buyrachy!$C$1,O6,,P6)),2,FALSE),FIND(""_"",VLOOKUP(" & Column_Letter(orcColDivisionNo + 6) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$D$1,O6,,P6),OFFSET(Buyrachy!$C$1,O6,,P6)),2,FALSE),1)-1),"""")"
        'BUYING SUB GROUP
            Range("R5") = "Buying Group Start"
            Range("R6") = "=O6+MATCH(" & Column_Letter(orcColDivisionNo + 6) & OracleHeadderRow + 2 & ",OFFSET(Buyrachy!$D$1,O6,,P6),0)-1"
            '13Mar20
            Range("S5") = "Buying Group Count"
            'Range("S6") = "=P6+COUNTIF(OFFSET(Buyrachy!$D$1,O6,,P6)," & Column_Letter(OrcColDivisionNo + 6) & OracleHeadderRow + 2 & ")"
            Range("S6") = "=COUNTIF(OFFSET(Buyrachy!$D$1,O6,,P6)," & Column_Letter(orcColDivisionNo + 6) & OracleHeadderRow + 2 & ")"
            Range("T5") = "Buying SubGroup ID"
            'Range("T6") = "=VLOOKUP(" & Column_Letter(OrcColDivisionNo + 7) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$F$1,R6,,S6),OFFSET(Buyrachy!$E$1,R6,,S6)),2,FALSE)"'Sprint 68
            Range("T6") = "=IFERROR(LEFT(VLOOKUP(" & Column_Letter(orcColDivisionNo + 7) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$F$1,R6,,S6),OFFSET(Buyrachy!$E$1,R6,,S6)),2,FALSE),FIND(""_"",VLOOKUP(" & Column_Letter(orcColDivisionNo + 7) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$F$1,R6,,S6),OFFSET(Buyrachy!$E$1,R6,,S6)),2,FALSE),1)-1),"""")"
        'BUYING SET
            Range("U5") = "Buying SubGroup Start"
            Range("U6") = "=R6+MATCH(" & Column_Letter(orcColDivisionNo + 7) & OracleHeadderRow + 2 & ",OFFSET(Buyrachy!$F$1,R6,,S6),0)-1"
            '13Mar20
            Range("V5") = "Buying SubGroup Count"
            'Range("V6") = "=S6+COUNTIF(OFFSET(Buyrachy!$F$1,R6,,S6)," & Column_Letter(OrcColDivisionNo + 7) & OracleHeadderRow + 2 & ")"
            Range("V6") = "=COUNTIF(OFFSET(Buyrachy!$F$1,R6,,S6)," & Column_Letter(orcColDivisionNo + 7) & OracleHeadderRow + 2 & ")"
            Range("W5") = "Buying Set ID"
            'Range("W6") = "=VLOOKUP(" & Column_Letter(OrcColDivisionNo + 8) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$H$1,U6,,V6),OFFSET(Buyrachy!$G$1,U6,,V6)),2,FALSE)"'Sprint 68
            Range("W6") = "=IFERROR(LEFT(VLOOKUP(" & Column_Letter(orcColDivisionNo + 8) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$H$1,U6,,V6),OFFSET(Buyrachy!$G$1,U6,,V6)),2,FALSE),FIND(""_"",VLOOKUP(" & Column_Letter(orcColDivisionNo + 8) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Buyrachy!$H$1,U6,,V6),OFFSET(Buyrachy!$G$1,U6,,V6)),2,FALSE),1)-1),"""")"
        'SUPPLIER SITE
            Range("X5") = "Supplier ID"
            Range("X6") = "=VLOOKUP(" & orcColSupplier & OracleHeadderRow + 2 & ",CHOOSE({1,2},RpasSuppliers!$B$1:$B$" & supMaxRows & ",RpasSuppliers!$A$1:$A$" & supMaxRows & "),2,FALSE)"
        'FACTORY
            Range("Y5") = "Supplier ID Start"
            Range("Y6") = "=MATCH(X6,SuppliersFactories!$A$1:$A$" & facMaxRows & ",0)-1"
            Range("Z5") = "Supplier ID Count"
            Range("Z6") = "=COUNTIF(SuppliersFactories!$A$1:$A$" & facMaxRows & ",X6)"
            Range("AA5") = "Factory ID"
            Range("AA6") = "=VLOOKUP(" & orcColFactoryUK & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(SuppliersFactories!$C$1,Y6,,Z6),OFFSET(SuppliersFactories!$B$1,Y6,,Z6)),2,FALSE)"
        'COLOUR GROUP
            Range("AB5") = "Colour Group ID"
            Range("AB6") = "=VLOOKUP(" & orcColColourGrp & OracleHeadderRow + 2 & ",CHOOSE({1,2},Diffs!$B$1:$B$" & diffsMaxRows & ",Diffs!$A$1:$A$" & diffsMaxRows & "),2,FALSE)"
        'COLOUR
            Range("AC5") = "Colour Group ID Start"
            Range("AC6") = "=MATCH(AB6,RpasDiffs!$E$1:$E$" & rpasDiffsMaxRows & ",0)-1"
            Range("AD5") = "Colour Group ID Count"
            Range("AD6") = "=COUNTIF(RpasDiffs!$E$1:$E$" & rpasDiffsMaxRows & ",AB6)"
            Range("AE5") = "Colour (Oracle) ID"
            Range("AE6") = "=VLOOKUP(" & orcColColour & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasDiffs!$B$1,AC6,,AD6),OFFSET(RpasDiffs!$A$1,AC6,,AD6)),2,FALSE)"
        'SIZE  GROUP
            Range("AF5") = "Size Data Start"
            Range("AF6") = "=MATCH(""Size"",Diffs!$D$1:$D$" & diffsMaxRows & ",0)-1"
            Range("AG5") = "Size Data Count"
            Range("AG6") = "=COUNTIF(Diffs!$D$1:$D$" & diffsMaxRows & ",""Size"")"

        'Cherry release
        'Range("AH6") = "=VLOOKUP(" & OrcColSizeGrp & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(RpasDiffs!$B$1,AF6,,AG6),OFFSET(RpasDiffs!$A$1,AF6,,AG6)),2,FALSE)"
            Range("AH5") = "Size Group ID"
            Range("AH6") = "=VLOOKUP(" & orcColSizeGrp & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(Diffs!$H$1,AF6,,AG6),OFFSET(Diffs!$G$1,AF6,,AG6)),2,FALSE)"

        'FACTORY
            Range("AI5") = "Brand ID"
            Range("AI6") = "=VLOOKUP(" & orcColBrand & OracleHeadderRow + 2 & ",CHOOSE({1,2},Brands!$B$1:$B$" & brndMaxRows & ",Brands!$A$1:$A$" & brndMaxRows & "),2,FALSE)"

    'Copypasting and Dragging Formula
        Range("A" & OracleHeadderRow + 2 & ":AI" & OracleHeadderRow + 2).Copy
        Range("A" & OracleHeadderRow + 2 & ":AI" & orcMaxRows).Select
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

        Range("A" & OracleHeadderRow + 2 & ":A" & orcMaxRows & ",D" & OracleHeadderRow + 2 & ":D" & orcMaxRows & ",G" & OracleHeadderRow + 2 & ":G" & orcMaxRows & ",J" & OracleHeadderRow + 2 & ":J" & orcMaxRows & ",M" & OracleHeadderRow + 2 & ":M" & orcMaxRows & ",N" & OracleHeadderRow + 2 & ":N" & orcMaxRows & ",Q" & OracleHeadderRow + 2 & ":Q" & orcMaxRows & ",T" & OracleHeadderRow + 2 & ":T" & orcMaxRows & ",W" & OracleHeadderRow + 2 & ":W" & orcMaxRows).Copy
        Range(orcColDivision & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        Range("AI" & OracleHeadderRow + 2 & ":AI" & orcMaxRows).Copy
        Range(orcColBrand & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        'Cherry release --- Changes End

    'Pasting Supplier IDs
        Range("X" & OracleHeadderRow + 2 & ":X" & orcMaxRows).Copy
        Range(orcColSupplier & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    'Pasting 3 Factory Columns
        Range("AA" & OracleHeadderRow + 2 & ":AA" & orcMaxRows).Copy
        Range(orcColFactoryUK & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Range(orcColFactoryEU & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Range(orcColFactoryUS & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    'Pasting Colour Group, Colour And Size Group Columns
        Range("AB" & OracleHeadderRow + 2 & ":AB" & orcMaxRows).Copy
        Range(orcColColourGrp & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        Range("AE" & OracleHeadderRow + 2 & ":AE" & orcMaxRows).Copy
        Range(orcColColour & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        Range("AH" & OracleHeadderRow + 2 & ":AH" & orcMaxRows).Copy
        Range(orcColSizeGrp & OracleHeadderRow + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    '==== Processing UDA Columns

 'Deleting few temporary columns
        'Columns("A:AD").Select
        Columns("A:AE").Select
        Selection.Delete Shift:=xlToLeft
        Range("A" & OracleHeadderRow).Select

        UDAColNo = 1

        While UDAColNo <= orcMaxCol
            If Left(Range(Column_Letter(UDAColNo) & OracleHeadderRow).Text, 3) = "UDA" Then

                    'Range("A6") = "=VLOOKUP(" & Column_Letter(UDAColNo) & OracleHeadderRow + 2 & ",CHOOSE({1,2},AttributesRaw!$B$1:$B$" & AttMaxRows & ",AttributesRaw!$A$1:$A$" & AttMaxRows & "),2,FALSE)"
                    'Range("B6") = "=MATCH(" & Column_Letter(UDAColNo) & OracleHeadderRow + 2 & ",AttributesRaw!$B$1:$B$" & AttMaxRows & ",0)-1"
                    'Range("C6") = "=COUNTIF(AttributesRaw!$B$1:$B$" & AttMaxRows & "," & Column_Letter(UDAColNo) & OracleHeadderRow + 2 & ")"

                    'Sprint 58
                    'Range("A6") = Mid(Range(Column_Letter(UDAColNo) & OracleHeadderRow), 8, 10)
                    UDAID_Temp = Mid(Range(Column_Letter(UDAColNo) & OracleHeadderRow), 8, 10) 'Updated on 4-Nov-2020
                    Range("A6") = "=IF(" & Column_Letter(UDAColNo + 1) & OracleHeadderRow + 2 & "= " & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & UDAID_Temp & ")" 'Updated on 4-Nov-2020

                    Range("B6") = "=MATCH(" & UDAID_Temp & ",AttributesRaw!$A$1:$A$" & attMaxRows & ",0)-1" 'updated by Salih 30-Mar-2021
                    Range("C6") = "=COUNTIF(AttributesRaw!$A$1:$A$" & attMaxRows & "," & UDAID_Temp & ")"   'updated by Salih 30-Mar-2021
                    Range("B6") = Range("B6").Text
                    Range("C6") = Range("C6").Text

                    'Checking if the field is dropdown to convert into IDs. If not dropdown then conversion will not happen
                    Range("D6") = "=VLOOKUP(" & UDAID_Temp & ",AttributesRaw!A:C,3,0)" 'updated by Salih 30-Mar-2021
                    If Range("D6").Text = "LV" Then
                        'Range("D6") = "=VLOOKUP(" & Column_Letter(UDAColNo + 1) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(AttributesRaw!$H$1,B6,,C6),OFFSET(AttributesRaw!$G$1,B6,,C6)),2,FALSE)"
                        'Range("D6") = "=VLOOKUP(" & Column_Letter(UDAColNo + 1) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(AttributesRaw!$H$1,B6,,C6),OFFSET(AttributesRaw!$G$1,B6,,C6)),2,FALSE)"
                         Range("D6") = "=if(" & Column_Letter(UDAColNo + 1) & OracleHeadderRow + 2 & "<>""""" & ",IFERROR(VLOOKUP(" & Column_Letter(UDAColNo + 1) & OracleHeadderRow + 2 & ",CHOOSE({1,2},OFFSET(AttributesRaw!$H$1,B6,,C6),OFFSET(AttributesRaw!$G$1,B6,,C6)),2,FALSE), ""Invalid UDA""),""#N/A"")" ' Added by Nirmala 02-Feb-2024
                    ElseIf Range("D6").Text = "DT" Then
                        Range("D6") = "=TEXT(" & Column_Letter(UDAColNo + 1) & OracleHeadderRow + 2 & ",""dd/mm/yyyy"")"
                    Else
                        Range("D6") = "=" & Column_Letter(UDAColNo + 1) & OracleHeadderRow + 2
                    End If

                'Copypasting and Dragging Formula
                    Range("A" & OracleHeadderRow + 2 & ":D" & OracleHeadderRow + 2).Copy
                    Range("A" & OracleHeadderRow + 2 & ":D" & orcMaxRows).Select
                    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                    Selection.Copy
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                    Range("A" & OracleHeadderRow).Select

                'Copypasting IDs into respective column
                    Range("A" & OracleHeadderRow + 2 & ":A" & orcMaxRows & ",D" & OracleHeadderRow + 2 & ":D" & orcMaxRows).Copy
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
    orcColBusinessModel = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "BUSINESS MODEL")
    orcColBuyingGrp = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "BUYING GROUP")
    orcColSizeGrp = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "SIZE GROUP")
    orcColSizeCurveIDUK = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "UK SIZE CURVE ID")
    orcColSizeCurveIDEU = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "EU SIZE CURVE ID")
    orcColSizeCurveIDUS = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "US SIZE CURVE ID")
    orcColManualSizeDistUK = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "UK MANUAL SIZE DIST")
    orcColManualSizeDistEU = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "EU MANUAL SIZE DIST")
    orcColManualSizeDistUS = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "US MANUAL SIZE DIST")
    orcColUKPoPriority = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "UK PO PRIORITY") 'Added by Manu on 30-Oct-2023
    orcColEUPoPriority = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "EU PO PRIORITY") 'Added by Manu on 30-Oct-2023
    orcColUSPoPriority = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "US PO PRIORITY") 'Added by Manu on 30-Oct-2023

'Copypasting and Dragging Formula
    Range("A" & OracleHeadderRow + 2) = "=IFERROR(VLOOKUP(" & orcColSizeCurveIDUK & OracleHeadderRow + 2 & "&""_""&" & orcColSizeGrp & OracleHeadderRow + 2 & "&""_""&" & orcColBuyingGrp & OracleHeadderRow + 2 & "&""_""&" & orcColBusinessModel & OracleHeadderRow + 2 & ",CascadingDropdowns!BH:BI,2,0),"""")"
    Range("B" & OracleHeadderRow + 2) = "=IFERROR(VLOOKUP(" & orcColSizeCurveIDEU & OracleHeadderRow + 2 & "&""_""&" & orcColSizeGrp & OracleHeadderRow + 2 & "&""_""&" & orcColBuyingGrp & OracleHeadderRow + 2 & "&""_""&" & orcColBusinessModel & OracleHeadderRow + 2 & ",CascadingDropdowns!BH:BI,2,0),"""")"
    Range("C" & OracleHeadderRow + 2) = "=IFERROR(VLOOKUP(" & orcColSizeCurveIDUS & OracleHeadderRow + 2 & "&""_""&" & orcColSizeGrp & OracleHeadderRow + 2 & "&""_""&" & orcColBuyingGrp & OracleHeadderRow + 2 & "&""_""&" & orcColBusinessModel & OracleHeadderRow + 2 & ",CascadingDropdowns!BH:BI,2,0),"""")"

    Range("A" & OracleHeadderRow + 2 & ":C" & OracleHeadderRow + 2).Copy
    Range("A" & OracleHeadderRow + 2 & ":C" & orcMaxRows).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

'Moving Size Curve IDs to respective columns
    For i = 1 To 3
        If i = 1 Then
            Range(orcColSizeCurveIDUK & OracleHeadderRow + 2) = "=IF(" & orcColManualSizeDistUK & OracleHeadderRow + 2 & " = """", IF(A6="""","""",1" & "&A6),"""")"
            Range(orcColSizeCurveIDUK & OracleHeadderRow + 2).Copy
            Range(orcColSizeCurveIDUK & OracleHeadderRow + 2 & ":" & orcColSizeCurveIDUK & orcMaxRows).Select
        ElseIf i = 2 Then
            Range(orcColSizeCurveIDEU & OracleHeadderRow + 2) = "=IF(" & orcColManualSizeDistEU & OracleHeadderRow + 2 & " = """", IF(B6="""","""",4" & "&B6),"""")"
            Range(orcColSizeCurveIDEU & OracleHeadderRow + 2).Copy
            Range(orcColSizeCurveIDEU & OracleHeadderRow + 2 & ":" & orcColSizeCurveIDEU & orcMaxRows).Select
        ElseIf i = 3 Then
            Range(orcColSizeCurveIDUS & OracleHeadderRow + 2) = "=IF(" & orcColManualSizeDistUS & OracleHeadderRow + 2 & " = """", IF(C6="""","""",3" & "&C6),"""")"
            Range(orcColSizeCurveIDUS & OracleHeadderRow + 2).Copy
            Range(orcColSizeCurveIDUS & OracleHeadderRow + 2 & ":" & orcColSizeCurveIDUS & orcMaxRows).Select
        End If
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    Next i

'Copypasting and Dragging Formula for CFA Columns: Added by Manu on 30-Oct-2023


    With ActiveWorkbook.Sheets("AttributesRaw")
    ' Count occurrences of value 3029 in column A
        poPriorityRow = Application.WorksheetFunction.Match(3029, .Columns("A"), 0)
        priorityFreq = Application.WorksheetFunction.CountIf(.Columns("A"), 3029)
    End With

    Range("A" & OracleHeadderRow + 2) = "=XLOOKUP(" & orcColUKPoPriority & OracleHeadderRow + 2 & ",AttributesRaw!$H$" & poPriorityRow & ":" & "$H$" & poPriorityRow + priorityFreq & ", AttributesRaw!$G$" & poPriorityRow & ":" & "$G$" & poPriorityRow + priorityFreq & ","""",0,1)"
    Range("B" & OracleHeadderRow + 2) = "=XLOOKUP(" & orcColEUPoPriority & OracleHeadderRow + 2 & ",AttributesRaw!$H$" & poPriorityRow & ":" & "$H$" & poPriorityRow + priorityFreq & ", AttributesRaw!$G$" & poPriorityRow & ":" & "$G$" & poPriorityRow + priorityFreq & ","""",0,1)"
    Range("C" & OracleHeadderRow + 2) = "=XLOOKUP(" & orcColUSPoPriority & OracleHeadderRow + 2 & ",AttributesRaw!$H$" & poPriorityRow & ":" & "$H$" & poPriorityRow + priorityFreq & ", AttributesRaw!$G$" & poPriorityRow & ":" & "$G$" & poPriorityRow + priorityFreq & ","""",0,1)"


    Range("A" & OracleHeadderRow + 2 & ":C" & OracleHeadderRow + 2).Copy
    Range("A" & OracleHeadderRow + 2 & ":C" & orcMaxRows).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

' Moving PO PRIORITY CFA IDs to respective columns: Added by Manu on 30-Oct-2023

For i = 1 To 3
        If i = 1 Then

            Range(orcColUKPoPriority & OracleHeadderRow + 2) = "=A6"
            Range(orcColUKPoPriority & OracleHeadderRow + 2).Copy
            Range(orcColUKPoPriority & OracleHeadderRow + 2 & ":" & orcColUKPoPriority & orcMaxRows).Select
        ElseIf i = 2 Then

            Range(orcColEUPoPriority & OracleHeadderRow + 2) = "=B6"
            Range(orcColEUPoPriority & OracleHeadderRow + 2).Copy
            Range(orcColEUPoPriority & OracleHeadderRow + 2 & ":" & orcColEUPoPriority & orcMaxRows).Select
        ElseIf i = 3 Then

            Range(orcColUSPoPriority & OracleHeadderRow + 2) = "=C6"
            Range(orcColUSPoPriority & OracleHeadderRow + 2).Copy
            Range(orcColUSPoPriority & OracleHeadderRow + 2 & ":" & orcColUSPoPriority & orcMaxRows).Select

        End If
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    Next i
    Range("A" & OracleHeadderRow).Select

    '****Converting shipping combination columns (SHIPPING POINT, SHIPPING METHOD & FREIGHT FORWARDER) data to IDs****
    'Finding the column position in the Oracle Upload Sheet
        orcUKSPCol = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "UK SHIPPING POINT")
        orcUKSMCol = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "UK SHIPPING METHOD")
        orcUKFFCol = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "UK FREIGHT FORWARDER")

        orcEUSPCol = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "EU SHIPPING POINT")
        orcEUSMCol = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "EU SHIPPING METHOD")
        orcEUFFCol = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "EU FREIGHT FORWARDER")

        orcUSSPCol = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "US SHIPPING POINT")
        orcUSSMCol = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "US SHIPPING METHOD")
        orcUSFFCol = Find_Specific_Columns(RPFileName, OracleSheet, OracleHeadderRow, "US FREIGHT FORWARDER")


    'Converting shipping combination texts to IDs for the UK Warehouse
        Call Covert_Shipping_Combinations_ToIDs(orcUKSPCol, orcUKSMCol, orcUKFFCol, DATSheetFreightMat, orcMaxRows)

    'Converting shipping combination texts to IDs for the UK Warehouse
        Call Covert_Shipping_Combinations_ToIDs(orcEUSPCol, orcEUSMCol, orcEUFFCol, DATSheetFreightMat, orcMaxRows)

    'Converting shipping combination texts to IDs for the UK Warehouse
        Call Covert_Shipping_Combinations_ToIDs(orcUSSPCol, orcUSSMCol, orcUSFFCol, DATSheetFreightMat, orcMaxRows)

    'Deleting remaining temporary columns
        Columns("A:D").Select
        Selection.Delete Shift:=xlToLeft
        Range("A" & OracleHeadderRow).Select
        Selection.Find(What:="*", LookAt:=xlPart).Activate
End Sub

Sub Covert_Shipping_Combinations_ToIDs(shipPointCol As String, shipMethodCol As String, freightForCol As String, DATSheetFreightMat As String, orcMaxRows As Long)

    'Converting shipping combination texts to IDs
        Range("A" & OracleHeadderRow + 2 & ":A" & orcMaxRows).formula = "=IFERROR(INDEX(" & DATSheetFreightMat & "!D:D,MATCH('" & OracleSheet & "'!" & shipPointCol & OracleHeadderRow + 2 & "," & DATSheetFreightMat & "!E:E,0)),"""")"
        Range("B" & OracleHeadderRow + 2 & ":B" & orcMaxRows).formula = "=IFERROR(LEFT(" & shipMethodCol & OracleHeadderRow + 2 & ",FIND(""-""," & shipMethodCol & OracleHeadderRow + 2 & ")-2),"""")"
        Range("C" & OracleHeadderRow + 2 & ":C" & orcMaxRows).formula = "=IFERROR(VLOOKUP(" & freightForCol & OracleHeadderRow + 2 & ",FreightMatrixRaw!J:L,3,0),"""")"
        Range("A" & OracleHeadderRow + 2 & ":A" & orcMaxRows).Copy
        Range(shipPointCol & OracleHeadderRow + 2).PasteSpecial xlPasteValues
        Range("B" & OracleHeadderRow + 2 & ":B" & orcMaxRows).Copy
        Range(shipMethodCol & OracleHeadderRow + 2).PasteSpecial xlPasteValues
        Range("C" & OracleHeadderRow + 2 & ":C" & orcMaxRows).Copy
        Range(freightForCol & OracleHeadderRow + 2).PasteSpecial xlPasteValues

End Sub
Function MacroVersionCheck_Oracle_TextToID_Conversion() As Boolean
    If Sheets(ManRefDataSheet).Range("AB18") = "Oracle_TextToID_Conversion_v107" Then MacroVersionCheck_Oracle_TextToID_Conversion = True
End Function
