#' Get dims parameters of the matrix yielded from the msoChart object
#'
#' Get the indices of rows and columnss involved in msoChart object. The function
#' will parse the series list yieled by \code{\link{getMsoChartSeries}}, and then
#' get the \strong{original} parameters out of the matrix. Please note that that
#' matrix may be different from that should be according to the actual chart.
#'
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param series series yielded from msoChart object by \code{getMsoChartSeries}
#'
#' @return A list comprising of 4 vectors: \describe{
#' \item{NamRows}{Row of Name part} \item{NamCols}{Col of Name part}
#' \item{ValRows}{Row of Values part} \item{ValCols}{Col of Values part}
#' \item{LblRows}{Row of XValues part} \item{lblCols}{Col of XValues part}}
#'
#' @export
#' @import cellranger
#' @seealso \code{\link{getMsoChartSeries}}
#' @examples
#' \dontrun{
#' ppt <- RDCOMClient::COMCreate("Powerpoint.Application")
#' pres <- ppt$Presentations()$Open(<some ppt file>)
#' slide <- pres$Slides(2)  # the 2nd slide
#' shape <- slide$Shapes(3)  # the 3rd shape
#' if (shape[['HasChart']] == -1) {  # if the shape contains msoChart
#'     chart <- shape[['Chart']]
#'     series <- getMsoChartSeries(chart)
#'     getMsoChartDims(series)
#' }
#' }
getMsoChartDims <- function(series){
    UseMethod(".getMsoChartDims")
}

#' @export
#' @rdname getMsoChartDims
get_msochart_dims <- getMsoChartDims

#' @export
.getMsoChartDims.MsoSeriesList <- function(series){
    dimParas <- list(NamRows=structure("name", rng=c('ul1', 'lr1')),
                     NamCols=structure("name", rng=c('ul2', 'lr2')),
                     ValRows=structure("value", rng=c('ul1', 'lr1')),
                     ValCols=structure("value", rng=c('ul2', 'lr2')),
                     LblRows=structure("label", rng=c('ul1', 'lr1')),
                     LblCols=structure("label", rng=c('ul2', 'lr2')),
                     WgtRows=structure("weight", rng=c('ul1', 'lr1')),
                     WgtCols=structure("weight", rng=c('ul2', 'lr2')))

    if (length(series) == 0) return(NULL)
    o <- lapply(dimParas, function(dimPara){
        unique(as.vector(unlist(sapply(series, function(lst){
            rng <- attr(dimPara, "rng")
            if (dimPara %in% names(lst)){
                return(eval(parse(text=paste0(
                    "c(",
                    paste(apply(lst[[dimPara]][, rng, drop=FALSE], 1,
                                function(row) paste(row, collapse=":")),
                          collapse=", "),
                    ")"))))
            }
        }))))
    })
    return(o)
}

#' @export
.getMsoChartDims.default <- .getMsoChartDims.MsoSeriesList
