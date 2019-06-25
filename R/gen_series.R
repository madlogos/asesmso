#' @import cellranger
#' @export
genSeries <- function(mtxIdx, mtx, series, chartType,
                      orient=c("vertical", "horizontal"), autoFill=FALSE, ...){
    # Generate a new series
    # mtxIdx and mtx are the table in source data sheet
    # series are data source mapping to the msoChart
    #
    # Args:
    #   mtxIdx: row/col index representing which row/col of the mtx
    #       to process.
    #   mtx: matrix
    #   series: series list (MsoSeriesList)
    #   chartType: Enum of the chart type.
    #   orient: series on columns (vertical) or rows (horizontal). If series is
    #       >0 in length, then get the orient from series[[1]].
    #   autoFill: auto fill the series values. Default FALSE.
    # Returns:
    #   A re-shaped series list (MsoSeriesList)

    # -------------------------Validation----------------------------
    stopifnot(length(series) == 0 || inherits(series, "MsoSeriesList"))
    stopifnot(is.matrix(mtx))
    # if series has values, get orient from series[[1]]$orient
    orient <- if (length(series) == 0) match.arg(orient) else
        series[[1]]$orient[[1]]
    stopifnot(is.numeric(mtxIdx))
    if (orient == "vertical") stopifnot(mtxIdx <= nrow(mtx)) else
        stopifnot(mtxIdx <= ncol(mtx))

    # mtx must be a core mtx
    if (! (any(!is.na(mtx[1, ])) && any(!is.na(mtx[nrow(mtx), ])) &&
           any(!is.na(mtx[, 1])) && any(!is.na(mtx[, ncol(mtx)]))))
        mtx <- getCoreDataMtx(mtx)

    # ------------------------- Classify -----------------------------
    ## 4 types: SeriesBrandNew, SeriesNew, SeriesExisting, SeriesNA
    if (length(series) == 0){
        class(mtxIdx) <- "SeriesBrandNew"  # totally new
    }else{
        dims <- getMsoChartDims(series)
        if (orient == "vertical"){
            if (mtxIdx %in% dims$ValCols){
                class(mtxIdx) <- "SeriesExisting"
            }else if (! mtxIdx %in% dims$WgtCols){
                class(mtxIdx) <- "SeriesNew"
            }else{
                class(mtxIdx) <- "SeriesNA"
            }
        }else{
            if (mtxIdx %in% dims$ValRows){
                class(mtxIdx) <- "SeriesExisting"
            }else if (! mtxIdx %in% dims$WgtRows){
                class(mtxIdx) <- "SeriesNew"
            }else{
                class(mtxIdx) <- "SeriesNA"
            }
        }
    }
    # ------------------------- use method -----------------------------
    UseMethod(".genSeries", mtxIdx)
}

#' @export
#' @rdname genSeries
gen_series <- genSeries

#' @import cellranger
#' @export
.genSeries.SeriesBrandNew <- function(
    mtxIdx, mtx, series, chartType, orient=c("vertical", "horizontal"),
    autoFill=autoFill, ...
){
    orient <- if (length(series) == 0) match.arg(orient) else
        series[[1]]$orient[[1]]
    isBubbleChart <- chartType %in%
        ENUM$xlChartType[c("xlBubble", "xlBubble3DEffect")]
    mtxcol <- c("ul1", "ul2", "lr1", "lr2", "sheet", "formula")

    args <- list(...)
    if ("defaultSheet" %in% names(args) && is.character(args$defaultSheet)){
        dftSht <- args$defaultSheet[[1]]
    }else{
        dftSht <- NA
    }

    if (orient == "vertical"){
        o <- list(name=matrix(c(1, mtxIdx, 1, mtxIdx, dftSht, NA), nrow=1,
                              dimnames=list(NULL, mtxcol)),
                  label=matrix(c(2, 1, nrow(mtx), 1, dftSht, NA), nrow=1,
                               dimnames=list(NULL, mtxcol)),
                  value=matrix(c(2, mtxIdx, nrow(mtx), mtxIdx, dftSht, NA),
                               nrow=1, dimnames=list(NULL, mtxcol)),
                  orient=orient)
        if (isBubbleChart)
            o[["weight"]] <- matrix(c(2, 2, nrow(mtx), 2, dftSht, NA), nrow=1,
                                    dimnames=list(NULL, mtxcol))
    }else{
        o <- list(name=matrix(c(mtxIdx, 1, mtxIdx, 1, dftSht, NA), nrow=1,
                              dimnames=list(NULL, mtxcol)),
                  label=matrix(c(1, 2, 1, ncol(mtx), dftSht, NA), nrow=1,
                               dimnames=list(NULL, mtxcol)),
                  value=matrix(c(mtxIdx, 2, mtxIdx, ncol(mtxIdx), dftSht, NA),
                               nrow=1, dimnames=list(NULL, mtxcol)),
                  orient=orient)
        if (isBubbleChart)
            o[["weight"]] <- matrix(c(2, 2, 2, ncol(mtx), dftSht, NA), nrow=1,
                                    dimnames=list(NULL, mtxcol))
    }
    o <- lapply(o, .refreshFormulaInSeriesMtx)
    return(structure(o, class="BrandNew"))
}

#' @import cellranger
#' @export
.genSeries.SeriesNew <- function(
    mtxIdx, mtx, series, chartType, orient=c("vertical", "horizontal"),
    autoFill=FALSE, ...
){
    orient <- if (length(series) == 0) match.arg(orient) else
        series[[1]]$orient[[1]]
    isBubbleChart <- chartType %in%
        ENUM$xlChartType[c("xlBubble", "xlBubble3DEffect")]

    o <- series[[length(series)]]  # copy last one

    if (autoFill) {
        o$value <- o$value[1, , drop=FALSE]
        o$label <- o$label[1, , drop=FALSE]
        if (isBubbleChart) o$weight <- o$weight[1, , drop=FALSE]
    }
    if (orient == "vertical"){
        o$name[, c("ul2", "lr2")] <- mtxIdx
        o$value[, c("ul2", "lr2")] <- mtxIdx
        if (isBubbleChart)
            o$weight[, c("ul2", "lr2")] <- mtxIdx
        if (autoFill){
            o$label[, "ul1"] <- o$value[, "ul1"] <- 2
            o$label[, "lr1"] <- o$value[, "lr1"] <- nrow(mtx)
            if (isBubbleChart){
                o$weight[, "ul1"] <- 2
                o$weight[, "lr1"] <- nrow(mtx)
            }
        }
    }else{
        o$name[, c("ul1", "lr1")] <- mtxIdx
        o$value[, c("ul1", "lr1")] <- mtxIdx
        if (isBubbleChart)
            o$weight[, c("ul1", "lr1")] <- mtxIdx
        if (autoFill){
            o$label[, "ul2"] <- o$value[, "ul2"] <- 2
            o$label[, "lr2"] <- o$value[, "lr2"] <- ncol(mtx)
            if (isBubbleChart){
                o$weight[, "ul2"] <- 2
                o$weight[, "lr2"] <- ncol(mtx)
            }
        }
    }
    o <- lapply(o, .refreshFormulaInSeriesMtx)
    return(structure(o, class="New"))
}

#' @import cellranger
#' @export
.genSeries.SeriesExisting <- function(
    mtxIdx, mtx, series, chartType, orient=c("vertical", "horizontal"),
    autoFill=FALSE, ...
){
    dims <- getMsoChartDims(series)

    orient <- if (length(series) == 0) match.arg(orient) else
        series[[1]]$orient[[1]]
    isBubbleChart <- chartType %in%
        ENUM$xlChartType[c("xlBubble", "xlBubble3DEffect")]

    o <- series[[which(dims[[
        ifelse(orient == "vertical", "ValCols", "ValRows")]] == mtxIdx)]]
    if (autoFill) {
        o$value <- o$value[1, , drop=FALSE]
        o$label <- o$label[1, , drop=FALSE]
        if (isBubbleChart) o$weight <- o$weight[1, , drop=FALSE]
    }
    if (orient == "vertical"){
        if (autoFill){
            o$label[, "ul1"] <- o$value[, "ul1"] <- 2
            o$label[, "lr1"] <- o$value[, "lr1"] <- nrow(mtx)
            if (isBubbleChart){
                o$weight[, "ul1"] <- 2
                o$weight[, "lr1"] <- nrow(mtx)
            }
        }
    }else{
        if (autoFill){
            o$label[, "ul2"] <- o$value[, "ul2"] <- 2
            o$label[, "lr2"] <- o$value[, "lr2"] <- ncol(mtx)
            if (isBubbleChart){
                o$weight[, "ul2"] <- 2
                o$weight[, "lr2"] <- ncol(mtx)
            }
        }
    }
    o <- lapply(o, .refreshFormulaInSeriesMtx)
    return(structure(o, class="Existing"))
}

#' @export
.genSeries.SeriesNA <- function(
    mtxIdx, mtx, series, chartType, orient=c("vertical", "horizontal"),
    autoFill=FALSE, ...
){
    return(NULL)
}

updateMsoChartSeries <- function(series, mtx, chartType,
                                 orient=c("vertical", "horizontal"),
                                 autoFill=FALSE, ...){
    # Add or modify series according to mtx
    # It is a wrapper of genSeries

    stopifnot(is.matrix(mtx))
    stopifnot(nrow(mtx) > 0 || ncol(mtx) > 0)
    stopifnot(inherits(series, "MsoSeriesList"))
    stopifnot(is.numeric(chartType))

    args <- list(...)  # you can put defaultSheet here, e.g., "Sheet1"

    if (length(series) == 0){  # brand new
        orient <- match.arg(orient)
        seriesNames <- 2:(if (orient=="vertical") ncol(mtx) else nrow(mtx))
        applyThru <- seriesNames
    }else{
        dims <- getMsoChartDims(series)
        orient <- series[[1]]$orient[[1]]
        if (orient == "vertical"){
            applyThru <- seq_len(ncol(mtx))[- c(dims$LblCols, dims$WgtCols)]
        }else{
            applyThru <- seq_len(nrow(mtx))[- c(dims$LblRows, dims$WgtRows)]
        }
    }
    o <- lapply(applyThru, genSeries, mtx=mtx, series=series,
                chartType=chartType, autoFill=autoFill, ...)
    names(o) <- seq_along(o)
    return(o)
}

update_msochart_series <- updateMsoChartSeries
