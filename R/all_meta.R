
getMsoChartAllMeta <- function(oChart, retainCoreAreaOnly=FALSE,
                            gapWidth=2, emptyString=c(NA, NaN)){
    # Work function: simultaneously get all the info from the msoChart
    # Wrapping getMsoChartSourceData and getMsoChartSeries
    # Args:
    #   refer to getMsoChartSourceData
    # Returns:
    #   list(chartType, sourceData, seriesList, dataSheet)
    if (! inherits(oChart, "COMIDispatch"))
        stop("oChart must be of COMIDispatch type. (pkg::RDCOMClient or rdcom)")
    if (inherits(try(oChart$Type(), silent=TRUE), "try-error") ||
        try(oChart$Parent()$HasChart() , silent) != -1)
        stop("oChart must be a msoChart.")

    # chart type
    enumChartType <- as.numeric(oChart$ChartType())
    stopifnot(enumChartType %in% ENUM$xlChartType)
    chartType <- ENUM$xlChartTyp[ENUM$xlChartType == enumChartType]
    if (length(chartType) == 0) chartType <- NULL

    # source data
    oChartdata <- oChart[["ChartData"]]
    invisible(oChartdata$Activate())
    oWb <- oChartdata[["Workbook"]]
    xlApp <- oWb$Application()
    xlApp[["Visible"]] <- FALSE
    oWs <- oWb$Worksheets(1)  # suppose the data is in the first worksheet
    dataSht <- oWs$Name()
    oUr <- oWs$UsedRange()
    sd <- sapply(oUr$Value(), function(lst) {
        lst <- sapply(lst, function(sublst)
            if (is.null(sublst)) sublst <- NA else sublst)
        unlist(lst)
    })
    if (retainCoreAreaOnly)
        sd <- getCoreDataMtx(sd, gapWidth=gapWidth,
                             emptyString=emptyString)

    # series
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
    series <- structure(o, class="MsoSeriesList")

    # return
    invisible(oWb$Close())
    invisible(xlApp$Quit())
    return(structure(list(chartType=chartType, sourceData=sd, seriesList=series,
                          dataSheet=dataSht),
           class="MsoChartMeta"))
}

get_msochart_allmeta <- getMsoChartAllMeta
