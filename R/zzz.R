#' asesmso: An R Analytics Toolkit on MS Office for ASES
#'
#' An analytics toolkit comprising of a series of work functions on MS Office operations.
#' The toolkit has a series of featured funcionalities specifically designed for ASES.
#'
#' @details This package is comprised of \cr \describe{
#'   \item{MSO Enums}{\code{\link{enum}()}}
#'   \item{PPT}{\code{\link{getPptElems}()}}
#'   \item{MSO Charts}{\code{\link{getMsoChartDims}()}, \code{\link{getMsoChartSeries}()},
#'     \code{\link{getMsoChartSourceData}()}, \code{\link{getMsoChartType}()},
#'     \code{\link{setMsoChartScale}()}, \code{\link{setMsoChartSeries}()},
#'     \code{\link{refresh_ppt_charts}()}}
#'   \item{Scales}{\code{\link{optimizeScaleRange}()}, \code{\link{optimizeScaleUnit}()}}
#' }
#'
#' @author \strong{Maintainer}: Yiying Wang, \email{wangy@@aetna.com}
#'
#' @importFrom magrittr %>%
#' @export %>%
#' @seealso \pkg{\link{aseskit}}
#' @docType package
#' @keywords internal
#' @name asesmso
NULL

#' @importFrom aseskit addRtoolsPath
.onLoad <- function(libname, pkgname="asesmso"){

    if (Sys.info()[['sysname']] == 'Windows'){
        Sys.setlocale('LC_CTYPE', 'Chs')
    }else{
        Sys.setlocale('LC_CTYPE', 'zh_CN.utf-8')
    }
    if (Sys.info()[['machine']] == "x64") if (Sys.getenv("JAVA_HOME") != "")
        Sys.setenv(JAVA_HOME="")

    addRtoolsPath()

    # pkgenv is a hidden env under pacakge:asesmso
    # -----------------------------------------------------------
    assign("pkgenv", new.env(), envir=parent.env(environment()))

    # ----------------------------------------------------------

    # options
    assign('op', options(), envir=pkgenv)
    options(stringsAsFactors=FALSE)

    pkgParam <- aseskit:::.getPkgPara(pkgname)
    toset <- !(names(pkgParam) %in% names(pkgenv$op))
    if (any(toset)) options(pkgParam[toset])
}

.onUnload <- function(libname, pkgname="asesmso"){
    op <- aseskit:::.resetPkgPara(pkgname)
    options(op)
}


.onAttach <- function(libname, pkgname="asesmso"){
    ver.warn <- ""
    latest.ver <- getOption(pkgname)$latest.version
    current.ver <- getOption(pkgname)$version
    if (!is.null(latest.ver) && !is.null(current.ver))
        if (latest.ver > current.ver)
            ver.warn <- paste0("\nThe most up-to-date version of ", pkgname, " is ",
			                   latest.ver, ". You are currently using ", current.ver)
    packageStartupMessage(paste("Welcome to", pkgname, current.ver,
                                 ver.warn))
}

