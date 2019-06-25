#' Get meta info of the shapes in a Powerpoint document
#'
#' You can get a data.frame of all the shapes across the ppt, which may help you
#' program scripts to paint charts.
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param files character vector of ppt filenames. Default NULL, and a GUI wizard
#' will launch to help you select the files.
#'
#' @return A list contaning data.frames with columns c("page", "shape", "shapeName",
#' "typeOfChart", "chartType", "chartIndex).
#'
#' @import RDCOMClient
#' @importFrom gWidgets2 gfile
#' @seealso You can refer to \code{ENUM$xlChartType} to see the Enum definitions
#' of the xlChartTypes.
#' @export
#'
#' @examples
#' \dontrun{
#' getPptElems(NULL)
#' }
getPptElems <- function(files=NULL){
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
                           multi=TRUE, toolkit=guiToolkit("tcltk"))
            files <- stringi::stri_conv(files, "CP936", "UTF-8")
        }
    }else{
        files <- unlist(files)
    }

    ppt <- COMCreate("Powerpoint.Application")
    meta <- lapply(files, function(pptfile){
        pres <- ppt$Presentations()$Open(pptfile)

        o <- lapply(seq_len(pres$Slides()$Count()), function(i){
            slide <- pres$Slides(i)
            t(sapply(seq_len(slide$Shapes()$Count()), function(j){
                shape <- slide$Shapes(j)
                if (shape[["HasChart"]] == -1){
                    chart <- shape$Chart()
                    structure(c(i, j, shape$Name(), chart$Type(),
                                chart$ChartType()),
                              names=c("page", "shape", "shapeName", "typeOfChart",
                                      "chartType"))
                }else{
                    structure(c(i, j, shape$Name(), NA, NA),
                              names=c("page", "shape", "shapeName", "typeOfChart",
                                      "chartType"))
                }
            }))
        })
        o <- as.data.frame(do.call("rbind", o), stringsAsFactors=FALSE)
        o[, "chartType"] <- sapply(as.numeric(o[, "chartType"]), function(x)
            if (is.na(x)) NA else names(ENUM$xlChartType)[which(ENUM$xlChartType==x)])
        o[, "page"] <- as.numeric(o[, "page"])
        o[, "shape"] <- as.numeric(o[, "shape"])
        o[, "typeOfChart"] <- as.numeric(o[, "typeOfChart"])

        o[, "chartIndex"] <- cumsum(!is.na(o[, "typeOfChart"]))
        o[is.na(o$typeOfChart), "chartIndex"] <- NA
        pres$Close()
        return(o)
    })
    names(meta) <- files
    ppt$Quit()
    return(meta)
}

#' @export
#' @rdname getPptElems
get_ppt_elems <- getPptElems
