#' Get the ChartType of a msoChart object
#'
#' If the msoChart is a pie chart, then this function will tell its formal chart
#' type: \cr
#' xlPie \cr
#' 5
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param oChart msoChart object (COMIDispatch class)
#'
#' @import RDCOMClient
#' @export
#' @return If the ChartType is valid, then a named numeric is returned. If no
#' valid ChartType can be found, returns \code{NULL}.
#' @examples
#' \dontrun{
#' ppt <- RDCOMClient::COMCreate("Powerpoint.Application")
#' pres <- ppt$Presentations()$Open(<some ppt file>)
#' slide <- pres$Slides(2)  # the 2nd slide
#' shape <- slide$Shapes(3)  # the 3rd shape
#' if (shape[['HasChart']] == -1) {  # if the shape contains msoChart
#'     chart <- shape[['Chart']]
#'     getMsoChartType(chart)
#' }}
getMsoChartType <- function(oChart){
    # Get the chart type of a msoChart object
    # Args:
    #   oChart: msoChart object (COMIDispatch class)
    # Returns:
    #   named num. If no match is found, then returns NULL.
    #
    if (! inherits(oChart, "COMIDispatch"))
        stop("oChart must be of COMIDispatch type. (pkg::RDCOMClient or rdcom)")
    if (inherits(try(oChart$Type(), silent=TRUE), "try-error") ||
        try(oChart$Parent()$HasChart() , silent) != -1)
        stop("oChart must be a msoChart.")
    enumChartType <- as.numeric(oChart$ChartType())
    stopifnot(enumChartType %in% ENUM$xlChartType)
    chartType <- ENUM$xlChartTyp[ENUM$xlChartType == enumChartType]
    if (length(chartType) == 0){
        return(NULL)
    }else{
        return(chartType)
    }
}

#' @export
#' @rdname getMsoChartType
get_msochart_type <- getMsoChartType
