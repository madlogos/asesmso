#' Auto suggest a range to beautify the axes of a masChart object
#'
#' Given a vector, automatically define the range to beautify the axes scale.
#' You can then use \code{\link{getOptScaleUnit}} to get the optimized major unit
#' of the scale.
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param x numeric vector
#'
#' @return numeric vector \code{c(min, max)}
#' @export
#' @import compiler
#' @examples
#' \dontrun{
#' optimizeScaleRange(c(3.1, 5.6))    # c(0, 6)
#' optimizeScaleRange(c(-4, -2))      # c(-4, 0)
#' optimizeScaleRange(c(-4.1, 15.6))  # c(-5, 20)
#' }
optimizeScaleRange <- function(x){
    stopifnot(is.numeric(x))
    x <- strsplit(sprintf("%E", range(x, na.rm=TRUE)), split="E")
    x <- lapply(x, as.numeric)

    funProperScale <- cmpfun(function(endpt, digit){
        # work function to return an optimized end point for the scale
        num <- abs(endpt)
        scalept <- if (num > 8) 10 else if (num > 7.5) 8 else
            if (num > 6) 7.5 else if (num > 5) 6 else if (num > 4) 5 else
                if (num > 3) 4 else if (num > 2) 3 else if (num > 1.5) 2 else
                    if (num > 1) 1.5 else 1
        return(10^digit * (if (endpt < 0) - scalept else scalept))
    })
    minScale <- if (x[[2]][1] >=0 && x[[1]][1] >= 0) 0 else
        funProperScale(x[[1]][1], x[[1]][2])
    maxScale <- if (x[[2]][1] < 0 && x[[1]][1] < 0) 0 else
        funProperScale(x[[2]][1], x[[2]][2])
    return(c(minScale, maxScale))
}

#' @export
#' @rdname optimizeScaleRange
optimize_scale_range <- optimizeScaleRange


getScaleRange <- function(series, dims, mtx, chartType){
    # Work function to get the optimized scale of an msoChart object
    # Args:
    #   series: the series list, yielded by getMsoChartSeries(oChart)
    #   dims: the dimOfChart list, yielded by getMsoChartDims(series)
    #   mtx: the matrix of the source data, yielded by getMsoChartSourceData(oChart)
    #   chartType: enum representing the chart type, yielded by chart$ChartType()
    # Returns:
    #   numeric vector c(min, max)
    if(! chartType %in% ENUM$xlChartType)
        stop("chartType must be %in% ", ENUM$xlChartType)
    if (length(series) == 0) {
        return(c(NA, NA))
    } else {
        stopifnot(is.list(series))
    }
    stopifnot(all(names(series[[1]]) %in% c(
        "name", "label", "value", "weight", "orient")))
    stopifnot(is.list(dims) && all(names(dims) %in% c(
        "NamRows", "NamCols", "ValRows", "ValCols", "LblRows", "LblCols",
        "WgtRows", "WgtCols")))
    stopifnot(is.matrix(mtx))
    chartType <- chartType[[1]]

    # shrink the matrix according to rownames and colnames, to get the core num part
    shrinkMtx <- mtx[as.character(dims$ValRows), as.character(dims$ValCols),
                     drop=FALSE]
    # convert the shrinked matrix to pure numeric
    shrinkMtx <- suppressWarnings(
        matrix(as.numeric(shrinkMtx), ncol=dim(shrinkMtx)[2]))

    if (chartType %in% ENUM$xlChartType[c(
        "xl3DAreaStacked100", "xl3DBarStacked100", "xl3DColumnStacked100",
        "xlAreaStacked100", "xlBarStacked100", "xlColumnStacked100",
        "xlConeBarStacked100", "xlConeColStacked100", "xlCylinderBarStacked100",
        "xlCylinderColStacked100", "xlLineMarkersStacked100", "xlLineStacked100",
        "xlPyramidBarStacked100", "xlPyramidColStacked100")]){
        scaleRange <- c(0, 1.05)
    }else if (chartType %in% ENUM$xlChartType[c(
        "xl3DAreaStacked", "xl3DBarStacked", "xl3DColumnStacked", "xlAreaStacked",
        "xlBarStacked", "xlColumnStacked", "xlConeBarStacked", "xlConeColStacked",
        "xlCylinderBarStacked", "xlCylinderColStacked", "xlLineMarkersStacked",
        "xlLineStacked", "xlPyramidBarStacked", "xlPyramidColStacked")]){
        scaleRange <- getOptScaleRange(
            if (series[[1]]$orient == "vertical") rowSums(shrinkMtx) else
                colSums(shrinkMtx))
    }else{
        scaleRange <- getOptScaleRange(as.vector(shrinkMtx))
    }

    return(scaleRange)
}

get_scale_range <- getScaleRange
