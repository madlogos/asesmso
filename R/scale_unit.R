#' Auto suggest a major unit number to beautify the axes of a masChart object
#'
#' Given a numeric range, automatically define a number that can serve as the
#' major unit to beautify the axes scale.
#' You are recommended to use \code{\link{getOptScaleRange}} beforehand to get
#' the optimized range of the scale.
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param ... you can either pass \code{c(min, max)} or \code{min}, \code{max}
#'
#' @return a numeric scalar
#' @export
#'
#' @examples
#' \dontrun{
#' # Both of the following expressions are OK
#' optimizeScaleUnit(0, 30) # 10
#' ## or
#' vec <- c(0, 30)
#' optimizeScaleUnit(vec)   # 10
#' }
optimizeScaleUnit <- function(...){
    args <- unlist(list(...), recursive=FALSE)
    if (length(args) > 1){

        args <- c(unlist(args[[1]])[1], unlist(args[[2]])[1])
        min <- min(args)
        max <- max(args)
    }else{
        args <- unlist(args)
        min <- min(args, na.rm=TRUE)
        max <- max(args, na.rm=TRUE)
    }
    if (! is.numeric(min) || ! is.numeric(max))
        stop("The arguments you input is invald. ",
             "It should be a numeric vector of two numeric scalers.")
    rng <- as.numeric(unlist(strsplit(sprintf("%E", max - min), split="E")))

    num <- round(rng[1] * 10, 0)
    digit <- rng[2]
    o <- if (num %in% c(100, 75, 50, 25, 10, 5)) num/50 * 10^digit else
        if (num %in% c(80, 40, 20, 12, 8, 4, 2, 1)) num/40 * 10^digit else
            if (num %in% c(90, 60, 30, 15, 6, 3)) num/30 * 10^digit else
                num/40 * 10^digit
    return(o)
}

#' @export
#' @rdname optimizeScaleUnit
optimize_scale_unit <- optimizeScaleUnit
