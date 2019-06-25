#' Retrieve source data matrix from the msoChart object
#'
#' The function calls \code{\link{RDCOMClient}} to retrieve the source data from
#' the embeded Worksheet of a msoChart object in the form of character matrix. It
#' basically creates COM objects step by step and fetch \code{Chart[["ChartData"]][[
#' "Workbook"]]$Worksheets(1)[["UsedRange"]][["Value"]]}.
#' You may then process the raw matrix using \code{\link{getCoreDataMtx}}.
#'
#' @note \itemize{
#' \item Currently \code{\link{RDCOMClient}} cannot retrieve non-ASCII characters.
#' \item By default, it only fetches the whole \emph{UsedRange} of the source data
#'  from the \strong{1st} worksheet.
#' \item It will open an MS Excel process to communicate via COM interface silently.
#' }
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param oChart msoChart object (COMIDispatch class)
#' @param retainCoreAreaOnly Logical, only obtain the upper left core area of the
#' matrix. Default FALSE. If TRUE, the function \code{\link{getCoreDataMtx}} will
#' be called with preset arguments (gapWidth=2, emptyString=c(NA, NaN)).
#' @param gapWidth numeric, width of the "gap" rows or columns. Default 2. If
#' \code{retainCoreAreaOnly} = TRUE, then take \code{gapWidth} into effect.
#' @param emptyString character vector, how to define the "gap".
#' Default \code{c(NA, NaN)}. If \code{retainCoreAreaOnly} = TRUE, then take
#' \code{emptyString} into effect.
#' @param ... Other arguments.
#'
#' @export
#' @import RDCOMClient
#'
#' @return A character matrix with character rownames and colnames.
#' Empty cells are converted to NA.
#' @seealso \code{\link{getCoreDataMtx}}  \code{\link{getCoreDataMtx}}
#' @examples
#' \dontrun{
#' ppt <- RDCOMClient::COMCreate("Powerpoint.Application")
#' pres <- ppt$Presentations()$Open(<some ppt file>)
#' slide <- pres$Slides(2)  # the 2nd slide
#' shape <- slide$Shapes(3)  # the 3rd shape
#' if (shape[['HasChart']] == -1) {  # if the shape contains msoChart
#'     chart <- shape[['Chart']]
#'     getMsoChartSourceData(chart)
#' }
#' }
getMsoChartSourceData <- function(oChart, retainCoreAreaOnly=FALSE,
                                  gapWidth=2, emptyString=c(NA, NaN), ...){
    if (! inherits(oChart, "COMIDispatch"))
        stop("oChart must be of COMIDispatch type. (pkg::RDCOMClient or rdcom)")
    if (inherits(try(oChart$Type(), silent=TRUE), "try-error") ||
        try(oChart$Parent()$HasChart() , silent) != -1)
        stop("oChart must be a msoChart.")

    oChartdata <- oChart[["ChartData"]]
    invisible(oChartdata$Activate())
    oWb <- oChartdata[["Workbook"]]
    xlApp <- oWb$Application()
    xlApp[["Visible"]] <- FALSE
    oWs <- oWb$Worksheets(1)  # suppose the data is in the first worksheet
    oUr <- oWs$UsedRange()
    o <- sapply(oUr$Value(), function(lst) {
        lst <- sapply(lst, function(sublst)
            if (is.null(sublst)) sublst <- NA else sublst)
        unlist(lst)
    })
    if (retainCoreAreaOnly) {
        o <- getCoreDataMtx(o, gapWidth=gapWidth, emptyString=emptyString)
    }else{
        rownames(o) <- seq_len(nrow(o))
        colnames(o) <- seq_len(ncol(o))
    }
    invisible(oWb$Close())
    invisible(xlApp$Quit())
    return(o)
}

#' @export
#' @rdname getMsoChartSourceData
get_msochart_sourcedata <- getMsoChartSourceData
