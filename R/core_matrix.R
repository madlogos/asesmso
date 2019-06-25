#' Shrink the matrix to its core area
#'
#' A matrix may contain separate data areas. This function will fetch the upper
#' left one as the core area, retaining the original row index and column index.
#' It is usually used after \code{\link{getMsoChartSourceData}}.
#'
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param mtx the raw matrix to process.
#' @param gapWidth numeric, width of the "gap" rows or columns. Default 2.
#' @param emptyString character vector, how to define the "gap".
#' Default \code{c(NA, NaN)}
#'
#' @return A shrinked matrix
#' @export
#' @seealso \code{\link{getMsoChartSourceData}}
#' @examples
#' \dontrun{
#' ## The source matrix
#' m <- matrix(c(NA, NA, NA, NA, NA, NA,
#'               NA, 0,  0,  NA, NA, 0,
#'               NA, 1,  2,  NA, NA, NA,
#'               NA, NA, NA, NA, NA, NA), byrow=TRUE, ncol=6)
#'
#' ## There are two areas, with a gap of 2 columns
#' getCoreDataMtx(m, gapWidth=2)  # yields the left upper 2*2 matrix
#'    2  3
#' 2  0  0
#' 3  1  2
#'
#' ## If the gapWidth is larger than 2, then the gap will be included in
#' ## the data area
#' getCoreDataMtx(m, gapWidth=3)  # yields:
#'    2  3  4  5  6
#' 2  0  0 NA NA  0
#' 3  1  2 NA NA NA
#'
#' ## If we include 0 in emptyString, then
#' getCoreDataMtx(m, gapWidth=2, emptyString=c(NA, NaN, 0))  # yields
#'    2  3
#' 3  1  2
#' }
getCoreDataMtx <- function(mtx, gapWidth=2, emptyString=c(NA, NaN)){
    # It uses the work function getMatrixMargins
    stopifnot(is.matrix(mtx))
    mtxMargin <- getMatrixMargins(mtx, gapWidth=gapWidth,
                                  emptyString=emptyString)
    seqRows <- seq(mtxMargin["ul1"], mtxMargin["lr1"])
    seqCols <- seq(mtxMargin["ul2"], mtxMargin["lr2"])
    mtx <- mtx[seqRows, seqCols]
    dim(mtx)=c(length(seqRows), length(seqCols))  # retain the structure of mtx
    rownames(mtx) <- seqRows
    colnames(mtx) <- seqCols
    return(mtx)
}

#' @export
#' @rdname getCoreDataMtx
get_core_matrix <- getCoreDataMtx

getMatrixMargins <- function(mtx, gapWidth=3, emptyString=c(NA, NaN)){
    # Detect the margin of a valid matrix
    # If there are side-by-side empty rows or columns ("gap"), the margin should
    # be located exactly at that place.
    # Args:
    #   mtx: a matrix
    #   gapWidth: how many continuous empty rows/columns should be deemed as a
    #        "gap"? default 3.
    #   emptyString: what is deamed as empty? default NA and NaN
    # Returns:
    #   A named vector "ul1" "ul2" "lr1" "lr2"
    #   (upper left row index, upper left column index,
    #    lower right row index, lower right column index)
    stopifnot(inherits(mtx, "matrix"))
    emptyColIdx <- seq_len(ncol(mtx))[apply(mtx, 2, function(v)
        all(v %in% emptyString))]
    emptyRowIdx <- seq_len(nrow(mtx))[apply(mtx, 1, function(v)
        all(v %in% emptyString))]

    getEdges <- function(emptyIdx, maxN){
        # if there are 2 side-by-side all-empty rows/cols, put edge there
        stopifnot((is.vector(emptyIdx) && is.numeric(emptyIdx)) ||
                      (is.numeric(maxN)))
        if (length(maxN) > 1) warning("Too many elements in maxN. ",
                                      "Only the first one will be used.")
        maxN <- maxN[[1]]
        if (maxN < 1) return(c(0, 0))

        validIdx <- seq_len(maxN)
        validIdx <- validIdx[! validIdx %in% emptyIdx]
        if (length(validIdx) == 0) return(c(0, 0))
        edgeMin <- min(validIdx)

        if (length(validIdx) > 1){
            runningDif <- c(validIdx[-1], NA) - validIdx
            bigGap <- which(runningDif >= gapWidth)
            edgeMax <- if (length(bigGap) > 0) validIdx[bigGap[1]] else
                max(validIdx)
        }else{
            edgeMax <- max(validIdx)
        }

        return(c(edgeMin, edgeMax))
    }

    validRowMargin <- getEdges(emptyRowIdx, nrow(mtx))
    validColMargin <- getEdges(emptyColIdx, ncol(mtx))

    return(c(ul1=validRowMargin[1], ul2=validColMargin[1],
             lr1=validRowMargin[2], lr2=validColMargin[2]))
}

get_matrix_margins <- getMatrixMargins
