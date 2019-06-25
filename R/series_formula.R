#' @import cellranger
refreshMsoChartFormula <- function(series){
    # refresh all the formula inside the msoSeries object
    # It is a wrapper of .refreshFormulaInSeriesMtx
    # Args:
    #   series: a series list (MsoSeriesList class)
    # Returns:
    #   A MsoSeriesList object with formula refreshed
    stopifnot(inherits(series, "MsoSeriesList"))
    if (length(series) == 0) return(series)
    series <- lapply(series, function(lst){
        lapply(lst, function(mtx){
            if (is.matrix(mtx)) .refreshFormulaInSeriesMtx(mtx) else mtx
        })
    })
    return(series)
}

refresh_msochart_formula <- refreshMsoChartFormula

#' @import cellranger
.refreshFormulaInSeriesMtx <- function(mtx){
    # refresh the formula in any matrix of a series list
    if (! is.matrix(mtx)) return(mtx)
    if (! all(colnames(mtx) %in% c(
        "ul1", "ul2", "lr1", "lr2", "sheet", "formula")))
        return(mtx)
    mtx[, "formula"] <- apply(mtx, 1, function(row){  # row by row
        if (is.na(row["sheet"])){
            if (any(is.na(row[c("ul1", "ul2", "lr1", "lr2")]))){
                NA
            }else if (row["ul1"] == row["lr1"] && row["ul2"] == row["lr2"]){
                # is a cell
                to_string(ra_ref(row["ul1"], TRUE, row["ul2"], TRUE),
                          strict=TRUE, fo="A1")
            }else{
                # is a range
                as.range(cell_limits(
                    ul=c(as.numeric(row["ul1"]), as.numeric(row["ul2"])),
                    lr=c(as.numeric(row["lr1"]), as.numeric(row["lr2"]))),
                    strict=TRUE, fo="A1")
            }
        }else{
            if (any(is.na(row[c("ul1", "ul2", "lr1", "lr2")]))){
                NA
            }else if (row["ul1"] == row["lr1"] && row["ul2"] == row["lr2"]){
                # is a cell
                to_string(ra_ref(row["ul1"], TRUE, row["ul2"], TRUE,
                                 sheet=row["sheet"]), strict=TRUE, fo="A1")
            }else{
                # is a range
                as.range(cell_limits(
                    ul=c(as.numeric(row["ul1"]), as.numeric(row["ul2"])),
                    lr=c(as.numeric(row["lr1"]), as.numeric(row["lr2"])),
                    sheet=row["sheet"]), strict=TRUE, fo="A1")
            }
        }
    })
    return(mtx)
}
