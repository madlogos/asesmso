#' Get the series parameters from a msoChart object
#'
#' Each series in a msoChart object can be represented as a formula:
#' "=SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$3,Sheet1!$B$2:$B$3,1)", comprising of
#' Name, XValues, Values (and/or BubbleSizes) and Index. This function can
#' retrieve the series parameters in list form.
#' @note This function will call MS Excel process to communicate with the mso
#' file via COM interface silently.
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param oChart msoChart object (COMIDispath class)
#' @param format (deprecated) "list"
#'
#' @return A list comprising of lists of name (Name), label (XValues),
#' value (Values) and/or weight (BubbleSizes) and orient (series on rows or columns),
#' e.g.: \cr
#' $`1``\cr
#' \describe{
#' \item{$ name}{\tabular{lllllll}{
#'  \tab "ul1" \tab "ul2" \tab "lr1" \tab "lr2" \tab "sheet" \tab "formula" \cr
#' [1,] \tab "1" \tab "2" \tab "1" \tab "2" \tab "Sheet1" \tab "Sheet1!$B$1"}
#' }
#' \item{$ label}{\tabular{lllllll}{
#'  \tab "ul1" \tab "ul2" \tab "lr1" \tab "lr2" \tab "sheet" \tab "formula" \cr
#' [1,] \tab "2" \tab "1" \tab "3" \tab "1" \tab "Sheet1" \tab "Sheet1!$A$2:$A$3" \cr
#' [2,] \tab "5" \tab "1" \tab "5" \tab "1" \tab "Sheet1" \tab "Sheet1!$A$5" }
#' }
#' \item{$ value}{\tabular{lllllll}{
#'  \tab "ul1" \tab "ul2" \tab "lr1" \tab "lr2" \tab "sheet" \tab "formula" \cr
#' [1,] \tab "2" \tab "2" \tab "3" \tab "2" \tab "Sheet1" \tab "Sheet1!$B$2:$B$3" \cr
#' [2,] \tab "5" \tab "2" \tab "5" \tab "2" \tab "Sheet1" \tab "Sheet1!$B$5" }
#' }
#' \item{$ orient}{[1] "vertical"}}
#'
#' @export
#' @import cellranger
#' @import RDCOMClient
#' @seealso \code{\link{RDCOMClient}}  \code{\link{cellranger}}
#' @references
#' Refer to \code{\link{cell_limits}} in \code{\link{cellranger}} to understand
#' the data structure ul1, ul2, lr1, lr2, sheet, ...
#' @examples
#' \dontrun{
#' ppt <- RDCOMClient::COMCreate("Powerpoint.Application")
#' pres <- ppt$Presentations()$Open(<some ppt file>)
#' slide <- pres$Slides(2)  # the 2nd slide
#' shape <- slide$Shapes(3)  # the 3rd shape
#' if (shape[['HasChart']] == -1) {  # if the shape contains msoChart
#'     chart <- shape[['Chart']]
#'     getMsoChartSeries(chart)
#' }
#' }
getMsoChartSeries <- function(oChart, format=c("list")){
    if (! inherits(oChart, "COMIDispatch"))
        stop("oChart must be of COMIDispatch type. (pkg::RDCOMClient or rdcom)")
    if (inherits(try(oChart$Type(), silent=TRUE), "try-error") ||
        try(oChart$Parent()$HasChart() , silent) != -1)
        stop("oChart must be a msoChart.")
    format <- match.arg(format)

    oChartData <- oChart$ChartData()
    invisible(oChartData$Activate())
    oWb <- oChartData[["Workbook"]]
    xlApp <- oWb$Application()
    xlApp[['Visible']] <- FALSE
    o <- lapply(seq_len(oChart$SeriesCollection()$Count()), function(i){
        # extract formula of the whole series
        formula <- sub("=SERIES\\((.+?)\\)$", "\\1",
                       oChart$SeriesCollection(i)$Formula())
        formula <- gsub("([^,\\(\\)]+)", "'\\1'", formula)
        formula <- gsub("\\(", "c\\(", formula)
        formula <- paste0("list(", formula, ")")
        formula <- gsub("([\\(,])([,\\)])", "\\1NA\\2", formula)

        formula <- eval(parse(text=formula))

        formula <- formula[sapply(formula, function(x) length(x) > 0)]
        names(formula) <- c("name", "label", "value", "index",
                            if (length(formula) > 4) "weight" else NULL)
        posIndex <- which(names(formula) == "index")
        # translate formula to cell_limits
        formula <- lapply(formula[- posIndex], function(v){
            cellLimits <- sapply(seq_along(v), function(i){
                if (grepl("\\$.+\\$|[A-Z]+[0-9]+|R[1-9]+C[1-9]+", v[i])){
                    o <- unlist(as.cell_limits(v[i]))
                    o <- c(o, c(formula=v[i]))
                }else{
                    o <- structure(
                        rep(NA, 6),
                        names=c("ul1", "ul2", "lr1", "lr2", "sheet", "formula"))
                }
            })
            return(t(cellLimits))
        })
        formula[["orient"]] <- if (formula$value[1, 'ul2'] == formula$value[1, 'lr2'])
            "vertical" else "horizontal"
        return(formula)
    })
    names(o) <- seq_along(o)

    invisible(oWb$Close())
    invisible(xlApp$Quit())

    if (format == "list"){
        return(structure(o, class="MsoSeriesList"))
    }
}

#' @export
#' @rdname getMsoChartSeries
get_msochart_series <- getMsoChartSeries
