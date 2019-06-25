#' Dataset: ENUM
#' 
#' Some of the common enumerations, majorly Micorsoft VBA Enums (for Word, Excel 
#' and PPT, etc) when applying mso-family functions in this package.
#' 
#' @name data:ENUM
#' @docType data
#' @author Yiying Wang \email{wangy@@aetna.com}
#' @details 
#' It contains a named list nested with named lists. \cr
#' \itemize{
#' \item Office core (23): \code{MsoAutoShape, MsoBackgroundStyleIndex, MsoBarPosition, 
#' MsoBaselineAlignment, MsoBulletType, MsoCharacterSet, MsoChartElementType, 
#' MsoDateTimeFormat, MsoFileType, MsoGradientStyle, MsoLineCapStyle, MsoLineStyle, 
#' MsoLineDashStyle, MsoOrientation, MsoParagraphAlignment, MsoPatternType, MsoSegmentType, 
#' MsoShapeType, MsoShapeStyleIndex, MsoTextDirection, MsoTextOrientation, 
#' MsoThemeColorIndex, MsoThemeColorSchemeInde}
#' \item Word (11): \code{WdConstants, WdSaveFormat, WdApplyQuickStyleSets, 
#' WdBuiltInProperty, WdColor, WdColorIndex, WdCellColor, WdBuiltinStyle, 
#' WdCaptionNumberStyle, WdCharacterCase, WdCountry}
#' \item Excel (59): \code{XlApplicationInternational, XlAxisCrosses, XlAxisGroup, 
#' XlAxisType, XlBackground, XlBarShape, XlBinsType, XlBorderWeight, XlCategoryLabelLevel, 
#' XlCategoryType, XlChartElementPosition, XlChartGallery, XlChartItem, XlChartPicturePlacement,
#' XlChartPictureType, XlChartSplitType, XlChartType, XlColorIndex, XlConstants, 
#' XlDataBarAxisPosition, XlDataBarBorderType, XlDataBarFillType, XlDataBarNegativeColorType, 
#' XlDataLabelPosition, XlDataLabelsType, XlDataSeriesDate, XlDataSeriesType, 
#' XlDeleteShiftDirection, XlDisplayBlanksAs, XlDisplayUnit, XlEndStyleCap, 
#' XlErrorBarDirection, XlErrorBarInclude, XlErrorBarType, XlFileAccess, XlFileFormat, 
#' XlHAlign, XlIcon, XlLegendPosition, XlLineStyle, XlMarkerStyle, XlOLEType, 
#' XlOrientation, XlPattern, XlPieSliceLocation, XlRgbColor, XlRowCol, XlScaleType, 
#' XlSeriesNameLevel, XlSheetType, XlTableStyleElementType, XlThemeColor, 
#' XlTickLabelOrientation, XlTickLabelPosition, XlTickMark, XlTimeUnit, XlTrendlineType, 
#' XlUnderlineStyle, XlVAlign}
#' \item PowerPoint (18): \code{PpBaselineAlignment, PpBorderType, PpBulletType, 
#' PpChangeCase, PpColorSchemeIndex, PpDateTimeFormat, PpEntryEffect, PpFollowColors, 
#' PpFrameColors, PpGuideOrientation, PpNumberedBulletStyle, PpParagraphAlignment, 
#' PpPlaceholderType, PpSaveAsFileType, PpSelectionType, PpSlideLayout, PpSlideSizeType, 
#' PpTextStyleType}
#' }
#' You can use \code{\link{enum}} function to extract the enum values.
#' 
#' @references 
#' You can read the Microsoft official API document \itemize{
#' \item Office core: \url{https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.core?view=office-pia}
#' \item Word: \url{https://docs.microsoft.com/en-us/office/vba/api/word(enumerations)} \cr
#' \item Excel: \url{https://docs.microsoft.com/en-us/office/vba/api/excel(enumerations)} \cr
#' \item PowerPoint: \url{https://docs.microsoft.com/en-us/office/vba/api/Powerpoint(enumerations)}
#' }
#' @examples 
#' \dontrun{
#' data(ENUM)
#' ENUM$MsoShapeType$msoChart  # returns
#' # [1] 3
#'
#' # or you can use
#' enum(msoChart)  # returns
#' # msoChart 
#' #        3 
#' }
#' @keywords data enum
#' 
NULL

