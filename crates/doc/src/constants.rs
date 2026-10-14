//! Enumeration constants of the word-processing object model (`wd...`), with the values the
//! public documentation gives.

const CONSTANTS: &[(&str, i32)] = &[
    // WdCollapseDirection
    ("wdCollapseStart", 1),
    ("wdCollapseEnd", 0),
    // WdUnits
    ("wdCharacter", 1),
    ("wdWord", 2),
    ("wdSentence", 3),
    ("wdParagraph", 4),
    ("wdLine", 5),
    ("wdStory", 6),
    ("wdScreen", 7),
    ("wdSection", 8),
    ("wdColumn", 9),
    ("wdRow", 10),
    ("wdWindow", 11),
    ("wdCell", 12),
    ("wdCharacterFormatting", 13),
    ("wdParagraphFormatting", 14),
    ("wdTable", 15),
    ("wdItem", 16),
    // WdMovementType
    ("wdMove", 0),
    ("wdExtend", 1),
    // WdStatistic
    ("wdStatisticWords", 0),
    ("wdStatisticLines", 1),
    ("wdStatisticPages", 2),
    ("wdStatisticCharacters", 3),
    ("wdStatisticParagraphs", 4),
    ("wdStatisticCharactersWithSpaces", 5),
    ("wdStatisticFarEastCharacters", 6),
    // WdParagraphAlignment
    ("wdAlignParagraphLeft", 0),
    ("wdAlignParagraphCenter", 1),
    ("wdAlignParagraphRight", 2),
    ("wdAlignParagraphJustify", 3),
    ("wdAlignParagraphDistribute", 4),
    // WdUnderline
    ("wdUnderlineNone", 0),
    ("wdUnderlineSingle", 1),
    ("wdUnderlineWords", 2),
    ("wdUnderlineDouble", 3),
    ("wdUnderlineDotted", 4),
    ("wdUnderlineThick", 6),
    ("wdUnderlineDash", 7),
    ("wdUnderlineWavy", 11),
    // WdConstants
    ("wdUndefined", 9_999_999),
    ("wdToggle", 9_999_998),
    ("wdForward", 1_073_741_823),
    ("wdBackward", -1_073_741_823),
    ("wdAutoPosition", 0),
    ("wdFirst", 1),
    ("wdCreatorCode", 1_297_307_460),
    // WdColor
    ("wdColorAutomatic", -16_777_216),
    ("wdColorBlack", 0),
    ("wdColorWhite", 16_777_215),
    ("wdColorRed", 255),
    ("wdColorGreen", 32_768),
    ("wdColorBrightGreen", 65_280),
    ("wdColorBlue", 16_711_680),
    ("wdColorYellow", 65_535),
    ("wdColorGray50", 8_421_504),
    ("wdColorDarkRed", 128),
    ("wdColorOrange", 26_367),
    // WdColorIndex
    ("wdAuto", 0),
    ("wdNoHighlight", 0),
    ("wdBlack", 1),
    ("wdBlue", 2),
    ("wdTurquoise", 3),
    ("wdBrightGreen", 4),
    ("wdPink", 5),
    ("wdRed", 6),
    ("wdYellow", 7),
    ("wdWhite", 8),
    // WdReplace
    ("wdReplaceNone", 0),
    ("wdReplaceOne", 1),
    ("wdReplaceAll", 2),
    // WdFindWrap
    ("wdFindStop", 0),
    ("wdFindContinue", 1),
    ("wdFindAsk", 2),
    // WdBuiltinStyle
    ("wdStyleNormal", -1),
    ("wdStyleHeading1", -2),
    ("wdStyleHeading2", -3),
    ("wdStyleHeading3", -4),
    ("wdStyleHeading4", -5),
    ("wdStyleHeading5", -6),
    ("wdStyleHeading6", -7),
    ("wdStyleHeading7", -8),
    ("wdStyleHeading8", -9),
    ("wdStyleHeading9", -10),
    ("wdStyleHeader", -32),
    ("wdStyleFooter", -33),
    ("wdStyleCaption", -35),
    ("wdStyleListBullet", -49),
    ("wdStyleListNumber", -50),
    ("wdStyleTitle", -63),
    ("wdStyleDefaultParagraphFont", -66),
    ("wdStyleBodyText", -67),
    ("wdStyleSubtitle", -75),
    ("wdStyleHyperlink", -86),
    ("wdStyleStrong", -88),
    ("wdStyleEmphasis", -89),
    ("wdStyleNoSpacing", -158),
    ("wdStyleListParagraph", -180),
    ("wdStyleQuote", -181),
    // WdStyleType
    ("wdStyleTypeParagraph", 1),
    ("wdStyleTypeCharacter", 2),
    ("wdStyleTypeTable", 3),
    ("wdStyleTypeList", 4),
    // WdSelectionType
    ("wdNoSelection", 0),
    ("wdSelectionIP", 1),
    ("wdSelectionNormal", 2),
    ("wdSelectionFrame", 3),
    ("wdSelectionColumn", 4),
    ("wdSelectionRow", 5),
    ("wdSelectionBlock", 6),
    ("wdSelectionInlineShape", 7),
    ("wdSelectionShape", 8),
    // WdSaveOptions
    ("wdDoNotSaveChanges", 0),
    ("wdSaveChanges", -1),
    ("wdPromptToSaveChanges", -2),
    // WdAlertLevel
    ("wdAlertsNone", 0),
    ("wdAlertsMessageBox", -2),
    ("wdAlertsAll", -1),
    // WdBreakType
    ("wdSectionBreakNextPage", 2),
    ("wdSectionBreakContinuous", 3),
    ("wdLineBreak", 6),
    ("wdPageBreak", 7),
    ("wdColumnBreak", 8),
    // WdStoryType
    ("wdMainTextStory", 1),
    // WdSaveFormat
    ("wdFormatDocument", 0),
    ("wdFormatTemplate", 1),
    ("wdFormatText", 2),
    ("wdFormatRTF", 6),
    ("wdFormatHTML", 8),
    ("wdFormatXMLDocument", 12),
    ("wdFormatXMLDocumentMacroEnabled", 13),
    ("wdFormatPDF", 17),
    ("wdFormatDocumentDefault", 16),
    // WdSortOrder / bookmarks
    ("wdSortByName", 0),
    ("wdSortByLocation", 1),
    // WdLineStyle
    ("wdLineStyleNone", 0),
    ("wdLineStyleSingle", 1),
    // WdGoToItem / WdGoToDirection
    ("wdGoToBookmark", -1),
    ("wdGoToPage", 1),
    ("wdGoToLine", 3),
    ("wdGoToAbsolute", 1),
    ("wdGoToNext", 2),
    ("wdGoToPrevious", 3),
    // WdInformation (selected)
    ("wdWithInTable", 12),
    ("wdActiveEndPageNumber", 3),
    ("wdFirstCharacterLineNumber", 10),
];

/// The value of a `wd...` constant (case-insensitive), if it is one.
pub fn constant(name: &str) -> Option<i32> {
    if name.len() < 3 || !name[..2].eq_ignore_ascii_case("wd") {
        return None;
    }
    CONSTANTS.iter().find(|(n, _)| n.eq_ignore_ascii_case(name)).map(|(_, v)| *v)
}

/// Built-in style names by their `wdStyle...` index.
pub(crate) fn builtin_style(index: i64) -> Option<&'static str> {
    Some(match index {
        -1 => "Normal",
        -2 => "Heading 1",
        -3 => "Heading 2",
        -4 => "Heading 3",
        -5 => "Heading 4",
        -6 => "Heading 5",
        -7 => "Heading 6",
        -8 => "Heading 7",
        -9 => "Heading 8",
        -10 => "Heading 9",
        -32 => "Header",
        -33 => "Footer",
        -35 => "Caption",
        -49 => "List Bullet",
        -50 => "List Number",
        -63 => "Title",
        -66 => "Default Paragraph Font",
        -67 => "Body Text",
        -75 => "Subtitle",
        -86 => "Hyperlink",
        -88 => "Strong",
        -89 => "Emphasis",
        -158 => "No Spacing",
        -180 => "List Paragraph",
        -181 => "Quote",
        _ => return None,
    })
}

/// Every built-in style name, for the Styles collection.
pub(crate) const BUILTIN_STYLES: &[i64] = &[
    -1, -2, -3, -4, -5, -6, -7, -8, -9, -10, -32, -33, -35, -49, -50, -63, -66, -67, -75, -86, -88, -89, -158, -180,
    -181,
];

/// True for character styles among the built-ins.
pub(crate) fn is_character_style(name: &str) -> bool {
    matches!(name, "Default Paragraph Font" | "Hyperlink" | "Strong" | "Emphasis")
}
