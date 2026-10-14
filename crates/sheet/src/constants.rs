//! The object model's documented enumeration constants (`xlUp`, `xlPasteValues`, ...).

/// Looks up an enumeration constant by name (case-insensitive).
pub fn constant(name: &str) -> Option<i32> {
    let lower = name.to_ascii_lowercase();
    CONSTANTS.iter().find(|(n, _)| n.eq_ignore_ascii_case(&lower)).map(|&(_, v)| v)
}

/// (name, value) for every constant the host resolves.
pub const CONSTANTS: &[(&str, i32)] = &[
    // XlDirection
    ("xlUp", -4162),
    ("xlDown", -4121),
    ("xlToLeft", -4159),
    ("xlToRight", -4161),
    // XlPasteType
    ("xlPasteAll", -4104),
    ("xlPasteAllExceptBorders", 7),
    ("xlPasteAllMergingConditionalFormats", 14),
    ("xlPasteAllUsingSourceTheme", 13),
    ("xlPasteColumnWidths", 8),
    ("xlPasteComments", -4144),
    ("xlPasteFormats", -4122),
    ("xlPasteFormulas", -4123),
    ("xlPasteFormulasAndNumberFormats", 11),
    ("xlPasteValidation", 6),
    ("xlPasteValues", -4163),
    ("xlPasteValuesAndNumberFormats", 12),
    // XlPasteSpecialOperation
    ("xlPasteSpecialOperationAdd", 2),
    ("xlPasteSpecialOperationDivide", 5),
    ("xlPasteSpecialOperationMultiply", 4),
    ("xlPasteSpecialOperationNone", -4142),
    ("xlPasteSpecialOperationSubtract", 3),
    // XlCalculation
    ("xlCalculationAutomatic", -4105),
    ("xlCalculationManual", -4135),
    ("xlCalculationSemiautomatic", 2),
    // XlCutCopyMode
    ("xlCopy", 1),
    ("xlCut", 2),
    // XlLineStyle
    ("xlContinuous", 1),
    ("xlDash", -4115),
    ("xlDashDot", 4),
    ("xlDashDotDot", 5),
    ("xlDot", -4118),
    ("xlDouble", -4119),
    ("xlLineStyleNone", -4142),
    ("xlSlantDashDot", 13),
    // XlBorderWeight
    ("xlHairline", 1),
    ("xlThin", 2),
    ("xlMedium", -4138),
    ("xlThick", 4),
    // XlBordersIndex
    ("xlDiagonalDown", 5),
    ("xlDiagonalUp", 6),
    ("xlEdgeLeft", 7),
    ("xlEdgeTop", 8),
    ("xlEdgeBottom", 9),
    ("xlEdgeRight", 10),
    ("xlInsideVertical", 11),
    ("xlInsideHorizontal", 12),
    // XlInsertShiftDirection / XlDeleteShiftDirection
    ("xlShiftDown", -4121),
    ("xlShiftToRight", -4161),
    ("xlShiftUp", -4162),
    ("xlShiftToLeft", -4159),
    // XlInsertFormatOrigin
    ("xlFormatFromLeftOrAbove", 0),
    ("xlFormatFromRightOrBelow", 1),
    // XlHAlign / XlVAlign
    ("xlGeneral", 1),
    ("xlHAlignGeneral", 1),
    ("xlLeft", -4131),
    ("xlHAlignLeft", -4131),
    ("xlCenter", -4108),
    ("xlHAlignCenter", -4108),
    ("xlRight", -4152),
    ("xlHAlignRight", -4152),
    ("xlFill", 5),
    ("xlHAlignFill", 5),
    ("xlJustify", -4130),
    ("xlHAlignJustify", -4130),
    ("xlCenterAcrossSelection", 7),
    ("xlHAlignCenterAcrossSelection", 7),
    ("xlDistributed", -4117),
    ("xlHAlignDistributed", -4117),
    ("xlTop", -4160),
    ("xlVAlignTop", -4160),
    ("xlBottom", -4107),
    ("xlVAlignBottom", -4107),
    ("xlVAlignCenter", -4108),
    // XlLookAt / XlFindLookIn / XlSearchOrder / XlSearchDirection
    ("xlWhole", 1),
    ("xlPart", 2),
    ("xlValues", -4163),
    ("xlFormulas", -4123),
    ("xlComments", -4144),
    ("xlByRows", 1),
    ("xlByColumns", 2),
    ("xlNext", 1),
    ("xlPrevious", 2),
    // XlCellType / XlSpecialCellsValue
    ("xlCellTypeAllFormatConditions", -4172),
    ("xlCellTypeAllValidation", -4174),
    ("xlCellTypeBlanks", 4),
    ("xlCellTypeComments", -4144),
    ("xlCellTypeConstants", 2),
    ("xlCellTypeFormulas", -4123),
    ("xlCellTypeLastCell", 11),
    ("xlCellTypeSameFormatConditions", -4173),
    ("xlCellTypeSameValidation", -4175),
    ("xlCellTypeVisible", 12),
    ("xlNumbers", 1),
    ("xlTextValues", 2),
    ("xlLogical", 4),
    ("xlErrors", 16),
    // XlYesNoGuess / XlSortOrder / XlSortOrientation
    ("xlGuess", 0),
    ("xlYes", 1),
    ("xlNo", 2),
    ("xlAscending", 1),
    ("xlDescending", 2),
    ("xlTopToBottom", 1),
    ("xlLeftToRight", 2),
    ("xlSortNormal", 0),
    ("xlSortOnValues", 0),
    // XlSheetVisibility / XlSheetType
    ("xlSheetVisible", -1),
    ("xlSheetHidden", 0),
    ("xlSheetVeryHidden", 2),
    ("xlWorksheet", -4167),
    ("xlChart", -4109),
    // Constants / XlColorIndex / XlPattern / XlUnderlineStyle
    ("xlNone", -4142),
    ("xlAutomatic", -4105),
    ("xlColorIndexAutomatic", -4105),
    ("xlColorIndexNone", -4142),
    ("xlSolid", 1),
    ("xlPatternSolid", 1),
    ("xlPatternNone", -4142),
    ("xlPatternAutomatic", -4105),
    ("xlGray16", 17),
    ("xlGray25", -4124),
    ("xlGray50", -4125),
    ("xlGray75", -4126),
    ("xlUnderlineStyleNone", -4142),
    ("xlUnderlineStyleSingle", 2),
    ("xlUnderlineStyleDouble", -4119),
    // XlReferenceStyle
    ("xlA1", 1),
    ("xlR1C1", -4150),
    // XlCVError
    ("xlErrDiv0", 2007),
    ("xlErrNA", 2042),
    ("xlErrName", 2029),
    ("xlErrNull", 2000),
    ("xlErrNum", 2036),
    ("xlErrRef", 2023),
    ("xlErrValue", 2015),
    // XlAutoFilterOperator
    ("xlAnd", 1),
    ("xlOr", 2),
    ("xlFilterValues", 7),
    ("xlTop10Items", 3),
    // XlFileFormat (commonly used)
    ("xlWorkbookDefault", 51),
    ("xlOpenXMLWorkbook", 51),
    ("xlOpenXMLWorkbookMacroEnabled", 52),
    ("xlCSV", 6),
    ("xlExcel8", 56),
    // XlWindowState / XlMousePointer / XlEnableCancelKey
    ("xlMaximized", -4137),
    ("xlMinimized", -4140),
    ("xlNormal", -4143),
    ("xlDefault", -4143),
    ("xlWait", 2),
    ("xlNorthwestArrow", 1),
    ("xlIBeam", 3),
    ("xlDisabled", 0),
    ("xlInterrupt", 1),
    ("xlErrorHandler", 2),
    // XlPageOrientation
    ("xlPortrait", 1),
    ("xlLandscape", 2),
    // XlFillWith / XlAutoFillType
    ("xlFillDefault", 0),
    ("xlFillCopy", 1),
    ("xlFillSeries", 2),
    ("xlFillFormats", 3),
    ("xlFillValues", 4),
    ("xlFillDays", 5),
    ("xlFillWeekdays", 6),
    ("xlFillMonths", 7),
    ("xlFillYears", 8),
    ("xlLinearTrend", 9),
    ("xlGrowthTrend", 10),
    // XlRowCol, XlDataSeriesType
    ("xlColumns", 2),
    ("xlRows", 1),
    ("xlLinear", -4132),
    ("xlGrowth", 2),
    ("xlChronological", 3),
    ("xlAutoFill", 4),
    // XlApplicationInternational
    ("xlDecimalSeparator", 3),
    ("xlThousandsSeparator", 4),
    ("xlListSeparator", 5),
    ("xlDateOrder", 32),
    ("xlCountryCode", 1),
    ("xlDateSeparator", 17),
    ("xlTimeSeparator", 18),
    // XlConsolidationFunction (subset)
    ("xlSum", -4157),
    ("xlCount", -4112),
    ("xlAverage", -4106),
    ("xlMax", -4136),
    ("xlMin", -4139),
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn names_are_unique() {
        let mut seen = std::collections::HashSet::new();
        for (n, _) in CONSTANTS {
            assert!(seen.insert(n.to_ascii_lowercase()), "duplicate {n}");
        }
        assert_eq!(constant("XLUP"), Some(-4162));
        assert_eq!(constant("nope"), None);
    }
}
