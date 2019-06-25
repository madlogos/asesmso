#' Optimize the scale range and major unit of the axes of a msoChart object
#'
#' It is a pipelined wrapper of \code{\link{getMsoChartDims}},
#' \code{\link{getMsoChartSeries}} and \code{\link{getOptScaleUnit}},
#' \code{\link{getOptScaleRange}} to deal with the source data matrix of the
#' mscChart object.
#' By default, the function directly sets the parameters MajorUnit, MinimumScale
#' and MaximumScale of the msoChart, and returns nothing about the optimized scale
#' range or major unit.
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param oChart msoChart object (COMIDispatch class)
#' @param changeScale Logical, change the MinimumScale and MaximumScale parameters
#' or not. Default \code{TRUE}.
#' @param returnValue Logical, return optimized parameters or not. Default
#' \code{FALSE}.
#' @param ... Other params.
#'
#' @return If returnValue is TRUE, then returns a list \code{list(MajorUnit,
#' ScaleRange)}. If FALSE, then returns nothing.
#' @export
#' @import RDCOMClient
#'
#' @seealso \code{\link{getMsoChartDims}}  \code{\link{getMsoChartSeries}}
#' \code{\link{getOptScaleUnit}}  \code{\link{getOptScaleRange}}
#' @examples
#' \dontrun{
#' ppt <- RDCOMClient::COMCreate("Powerpoint.Application")
#' pres <- ppt$Presentations()$Open(<some ppt file>)
#' slide <- pres$Slides(2)  # the 2nd slide
#' shape <- slide$Shapes(3)  # the 3rd shape
#' if (shape[['HasChart']] == -1) {  # if the shape contains msoChart
#'     chart <- shape[['Chart']]
#'     setMsoChartScale(chart)
#' }
#' }
setMsoChartScale <- function(oChart, changeScale=TRUE, returnValue=FALSE, ...) {
    if (! inherits(oChart, "COMIDispatch"))
        stop("oChart must be of COMIDispatch type. (pkg::RDCOMClient or rdcom)")
    if (inherits(try(oChart$Type(), silent=TRUE), "try-error") ||
        try(oChart$Parent()$HasChart() , silent) != -1)
        stop("oChart must be a msoChart.")

    # Pie and doughnut charts do not have axes
    if (! oChart$ChartType() %in% ENUM$xlChartType[c(
        "xl3DPie", "xl3DPieExploded", "xlBarOfPie", "xlDoughnut",
        "xlDoughnutExploded", "xlPie", "xlPieExploded", "xlPieOfPie")]){

        args <- list(...)
        if (length(args) > 0) {
            if ("chartMeta" %in% names(args) &&
                any(class(args$chartMeta) %in% "MsoChartMeta")){
                chartMeta <- args[['chartMeta']]
                chartType <- chartMeta[["chartType"]]
                series <- chartMeta[["seriesList"]]
                dimsOfChart <- getMsoChartDims(series)
                mtx <- chartMeta[["sourceData"]]
            }
        }else{
            chartType <- oChart$ChartType()
            series <- getMsoChartSeries(oChart)
            dimsOfChart <- getMsoChartDims(series)
            mtx <- getMsoChartSourceData(oChart)
        }

        scaleRange <- getScaleRange(series=series, dims=dimsOfChart, mtx=mtx,
                                    chartType=chartType)
        majorUnit <- getOptScaleUnit(scaleRange)

        if (changeScale){
            oAxes <- oChart$Axes(ENUM$xlAxisType['xlValue'])  # 2: xlValue
            if (!is.na(scaleRange[1])) oAxes[["MinimumScale"]] <- scaleRange[1]
            if (!is.na(scaleRange[2])) oAxes[["MaximumScale"]] <- scaleRange[2]
            if (!is.na(majorUnit)) oAxes[["MajorUnit"]] <- majorUnit
        }
        if (returnValue)
            return(list(MajorUnit=majorUnit, ScaleRange=scaleRange))
    }
}

#' @export
#' @rdname setMsoChartSeries
set_msochart_scale <- setMsoChartScale
