#' Automatically refresh msoChart in a PPT presentation
#'
#' It calls \code{\link{setMsoChartSeries}} and \code{\link{setMsoChartScale}} to
#' match source data with chart series and axes scales.
#' @author Yiying Wang, \email{wangy@@aetna.com}
#'
#' @param files character vector of ppt files. If NULL, a GUI wizard will launch.
#' @param orient character, "vertical" or "horizontal". How the series is generated.
#' @param autoFill logical, whether fill the series with the entire row/col. Default
#' FALSE
#' @param save Logical, whether save the file after processing. Default TRUE.
#' @param ... other arguments, e.g., \code{gapWidth} and \code{emptyString} for
#' \code{\link{getMsoChartSourceData}}.
#'
#' @return Nothing
#' @import RDCOMClient
#' @importFrom stringi stri_conv
#' @importFrom gWidgets2 gfile
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
#'     refreshPptCharts(NULL, orient="vertical", autoFill=TRUE)
#' }
#' }
refreshPptCharts <- function(files=NULL, orient=c("vertical", "horizontal"),
                                 autoFill=FALSE, save=TRUE, ...){
    if (is.null(files)) {
        if (Sys.info()['sysname'] == "Windows"){
            files <- invisible(choose.files(
                paste0(getOption("init.dir"), "/*.*"),
                caption="Select .ppt files...",
                filters=rbind(matrix(c("ppt files(.ppt)", "*.ppt?;*.PPT?"),
                                     nrow=1), Filters["All",])))
        }else{
            files <- gfile(text="Select .ppt files...", type="open",
                           initial.dir=getOption("init.dir"),
                           filter=c('ppt files'='ppt?'),
                           multi=TRUE, toolkit=guiToolkit(getOption("guiToolkit")))
            files <- stri_conv(files, "CP936", "UTF-8")
        }
    }else{
        files <- unlist(files)
    }
    files <- files[grepl("\\.ppt[xmb]?$", tolower(files))]
    if (length(files) == 0) stop("There is no valid ppt file assigned.")

    if (! require(RDCOMClient))
        stop("You need to explicitly library(RDCOMClient) before library(aseshms).")

    ppt <- COMCreate("Powerpoint.Application")
    invisible(sapply(files, function(pptfile){
        pres <- ppt$Presentations()$Open(pptfile)
        invisible(sapply(seq_len(pres$Slides()$Count()), function(i){
            slide <- pres$Slides(i)
            invisible(sapply(seq_len(slide$Shapes()$Count()), function(j){
                shape <- slide$Shapes(j)
                if (shape[["HasChart"]] == -1){
                    # browser()
                    chart <- shape$Chart()
                    chartMeta <- getMsoChartAllMeta(chart,
                                                    retainCoreAreaOnly=TRUE)
                    setMsoChartSeries(chart, orient=orient, autoFill=autoFill,
                                      chartMeta=chartMeta)
                    chartMeta$seriesList <- getMsoChartSeries(chart)
                    setMsoChartScale(chart, chartMeta=chartMeta)
                }
            }))
        }))
        if (save){
            pres$Save()
            invisible(pres$Close())
            invisible(ppt$Quit())
        }
        message("All the charts in  '", pptfile,
                "' have been successfully refreshed. \nPlease double check.")
    }))
}

#' @export
#' @rdname refreshPptCharts
refresh_ppt_charts <- refreshPptCharts

