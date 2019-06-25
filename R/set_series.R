#' Reset an msoChart object to match the source data and series
#'
#' It fetches meta data from a msoChart object and then reset the series one by
#' one. The msoChart will be visually modified.
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param oChart msoChart object
#' @param orient character, "vertical" or "horizontal". How the series is generated.
#' @param autoFill logical, whether fill the series with the entire row/col. Default
#' FALSE
#' @param ... other arguments, e.g., \code{gapWidth} and \code{emptyString} for
#' \code{\link{getMsoChartSourceData}}.
#'
#' @return Nothing
#' @export
#'
#' @examples
#' \dontrun{
#' ppt <- RDCOMClient::COMCreate("Powerpoint.Application")
#' pres <- ppt$Presentations()$Open(<some ppt file>)
#' slide <- pres$Slides(2)  # the 2nd slide
#' shape <- slide$Shapes(3)  # the 3rd shape
#' if (shape[['HasChart']] == -1) {  # if the shape contains msoChart
#'     chart <- shape[['Chart']]
#'     setMsoChartSeries(chart, orient="vertical", autoFill=TRUE)
#' }}
setMsoChartSeries <- function(oChart, orient=c("vertical", "horizontal"),
                              autoFill=FALSE, ...){
    if (! inherits(oChart, "COMIDispatch"))
        stop("oChart must be of COMIDispatch type. (pkg::RDCOMClient or rdcom)")
    if (inherits(try(oChart$Type(), silent=TRUE), "try-error") ||
        try(oChart$Parent()$HasChart() , silent) != -1)
        stop("oChart must be a msoChart.")
    seriesElem <- structure(c("name", "label", "value", "weight"),
                            names=c("Name", "XValues", "Values", "BubbleSizes"))

    args <- list(...)
    if (length(args) > 0) {
        if ("chartMeta" %in% names(args) &&
            any(class(args$chartMeta) %in% "MsoChartMeta")){
            chartMeta <- args[['chartMeta']]
        }
    }else{
        chartMeta <- getMsoChartAllMeta(oChart, retainCoreAreaOnly=TRUE)
    }
    chartType <- chartMeta$chartType
    series <- chartMeta$seriesList
    mtx <- chartMeta$sourceData
    dims <- getMsoChartDims(series)
    dftSht <- chartMeta$dataSheet

    if (length(series) == 0){
        orient <- match.arg(orient)
    }else{
        orient <- series[[1]]$orient[[1]]
    }

    # which to remove
    whichToRemove <- if (orient == "vertical"){
        which(! dims[['ValCols']] %in% seq_len(ncol(mtx)))
    }else{
        which(! dims[['ValRows']] %in% seq_len(nrow(mtx)))
    }
    if (length(whichToRemove) > 0)
        invisible(sapply(whichToRemove, function(i)
            oChart$SeriesCollection(i)$Delete()))

    # which to add or modify
    series <- structure(updateMsoChartSeries(
        series, mtx, chartType, orient=orient, autoFill=autoFill,
        defaultSheet=dftSht), class="MsoSeriesList")
    formula <- renderMsoChartSeries(series)
    fmlClass <- sapply(formula, class)
    iExisting <-  cumsum(fmlClass=="Existing")
    invisible(sapply(seq_along(formula), function(i){
        if (class(formula[[i]]) %in% c("New", "BrandNew")){
            ns <- oChart$SeriesCollection()$NewSeries()
        }else if (class(formula[[i]]) %in% c("Existing")){
            ns <- oChart$SeriesCollection(iExisting[i])
        }

        elemName <- intersect(names(formula[[i]]), seriesElem)
        invisible(sapply(elemName, function(elem){
            lblElem <- names(seriesElem)[seriesElem == elem]
            if (is.na(formula[[i]][elem])){
                ns[[lblElem]] <- ""
            }else{
                ns[[lblElem]] <- if (elem == "name")
                    paste0("=", formula[[i]][elem]) else formula[[i]][elem]
            }
        }))
    }))
    # oChart$Refresh()
    oChartdata <- oChart[["ChartData"]]
    invisible(oChartdata$Activate())
    oWb <- oChartdata[["Workbook"]]
    xlApp <- oWb$Application()
    xlApp[["Visible"]] <- FALSE
    invisible(oWb$Close())
    invisible(xlApp$Quit())
}

#' @export
#' @rdname setMsoChartSeries
set_msochart_series <- setMsoChartSeries
