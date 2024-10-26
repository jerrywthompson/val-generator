Attribute VB_Name = "GlobalVars"
Public Const gblGeneratorVersion = "2.6" 'end of row

Public Const startHdrFldRow = 3

Public updateCode As Boolean

Public userInputErrChk As Boolean

Public Const baseFieldRangeStart = "A"
Public Const baseFieldRangeRuleStart = "D"
Public Const baseFieldRangeEnd = "Q"

Public Const filterFieldRangeStart = "A"
Public Const filterFieldRangeRuleStart = "E"
Public Const filterFieldRangeEnd = "R"

Public Const concatFieldRangeStart = "A"
Public Const concatFieldRangeRuleStart = "F"
Public Const concatFieldRangeEnd = "T"

'set column locations for Base Fields sheet
Public Const bBeginIndexCol = 1
Public Const bEndIndexCol = 2
Public Const bFieldNameCol = 3
Public Const bDataTypeCol = 4
Public Const bDataTypeLvlCol = 5
Public Const bRequiredCol = 6
Public Const bRequiredLvlCol = 7
Public Const bUniqueCol = 8
Public Const bUniqueLvlCol = 9
Public Const bFormatCol = 10
Public Const bEnumValuesCol = 11
Public Const bRegExCol = 12
Public Const bFormatLvlCol = 13
Public Const bLowRangeCol = 14
Public Const bHighRangeCol = 15
Public Const bRangeLvlCol = 16
Public Const bHistoCol = 17

'set column locations for Filtered Fields sheet
Public Const fComputedFieldNameCol = 1
Public Const fFilterFieldCol = 2
Public Const fFilterCol = 3
Public Const fFilterValueCol = 4
Public Const fDataTypeCol = 5
Public Const fDataTypeLvlCol = 6
Public Const fRequiredCol = 7
Public Const fRequiredLvlCol = 8
Public Const fUniqueCol = 9
Public Const fUniqueLvlCol = 10
Public Const fFormatCol = 11
Public Const fEnumValuesCol = 12
Public Const fRegExCol = 13
Public Const fFormatLvlCol = 14
Public Const fLowRangeCol = 15
Public Const fHighRangeCol = 16
Public Const fRangeLvlCol = 17
Public Const fHistoCol = 18

'set column locations for Concat Fields sheet
Public Const cFieldName1Col = 1
Public Const cFieldName2Col = 2
Public Const cFieldName3Col = 3
Public Const cFieldName4Col = 4
Public Const cFieldName5Col = 5
Public Const cOutputDelimiterCol = 6
Public Const cDataTypeCol = 7
Public Const cDataTypeLvlCol = 8
Public Const cRequiredCol = 9
Public Const cRequiredLvlCol = 10
Public Const cUniqueCol = 11
Public Const cUniqueLvlCol = 12
Public Const cFormatCol = 13
Public Const cEnumValuesCol = 14
Public Const cRegExCol = 15
Public Const cFormatLvlCol = 16
Public Const cLowRangeCol = 17
Public Const cHighRangeCol = 18
Public Const cRangeLvlCol = 19
Public Const cHistoCol = 20

'set column locations for Coded Fields sheet
Public Const cdConfigFieldNameCol = 1
Public Const cdConceptCol = 2
Public Const cdCodeIDCol = 3
Public Const cdCodeSysIDCol = 4
Public Const cdCodeDisplayCol = 5
Public Const cdDefaultCodeSysCol = 6

Public Const delimiterValue = "PIPE;TAB;CSV;FULLCSV"
Public Const vendorValue = "allscripts;epic;meditech"

