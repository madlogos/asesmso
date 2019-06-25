
renderMsoChartSeries <- function(series){
    # Build series formula based on series (MsoSeriesList class)
    # wrapper of .renderMsoChartSeries
    stopifnot(inherits(series, "MsoSeriesList"))
    lapply(series, .renderMsoChartSeries)
}

render_msochart_series <- renderMsoChartSeries

.renderMsoChartSeries <- function(series){
    # Return list(Name, XValues, Values, BubbleSize) based on series fragment
    stopifnot(all(names(series) %in% c(
        "name", "label", "value", "weight", "orient")))
    seriesElem <- structure(c("name", "label", "value", "weight"),
                            names=c("Name", "XValues", "Values", "BubbleSizes"))
    o <- sapply(intersect(names(series), seriesElem), function(elem){
        if (nrow(series[[elem]] > 1)){
            paste(series[[elem]][, 'formula'], collapse=", ")
        }else{
            series[[elem]][, 'formula', drop=TRUE]
        }
    })
    names(o) <- intersect(names(series), seriesElem)
    return(structure(o, class=class(series)))
}



