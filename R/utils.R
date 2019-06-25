# ------------get enum--------------

#' Get enumeration value(s) based on the enum name(s)
#'
#' There is a builtin ENUM data set storing enumerations that are quite handy when
#' applying mso-family functions. 
#' @author Yiying Wang, \email{wangy@@aetna.com}
#' @param enum_name vector, can be \describe{
#' \item{scalar character}{e.g., "msoChart". In this case, you can pass other characters
#' in \code{...}}
#' \item{vector character}{e.g, c("msoChart", "msoTable") or list("msoChart", 
#' "msoTable"). In this case, \code{...} is omitted.}
#' \item{scalar symbol}{e.g., msoChart. In this case, you can pass other symbols
#' in \code{...}}
#' \item{vector symbol}{e.g, c(msoChart, msoTable) or list(msoChart, 
#' msoTable). In this case, \code{...} is omitted.}
#' }
#' @param ... When \code{enum_name} is scalar, you can pass the rest of the enum_names
#' here.
#' @param family character vector of the enumeration families, e.g., "MsoFileType".
#' You can use \code{names(ENUM)} to check the full list. When it is defined, the
#' function will only search enumerations within the families specified. Default NULL,
#' i.e. the function will search the whole ENUM data set. You can only input the
#' initial letters of the families.
#' @param unlist logical, whether unlist the result. Default TRUE.
#' @param bare logical, whether only return bare enumeration numbers. Default FALSE.
#'
#' @return \describe{
#' \item{If \code{unlist} = TRUE}{return a vector, with a vector of 'family' attribute 
#' on the vector (of same length as the result)}
#' \item{If \code{unlist} = FALSE}{return a list, with a 'family' attribute on each element}
#' \item{If \code{enum_name} is NULL or NA}{return all the enumerations that match}
#' \item{If nothing is found}{return NA}
#' }
#' @export
#'
#' @seealso \code{\link{data:ENUM}}
#' @examples
#' \dontrun{
#' enum(msoChart)  # equivalent to enum("msoChart")
#' # msoChart 
#' #        3 
#' # attr(,"family")
#' # [1] "msoShapeType"
#' 
#' enum(msoChart, unlist=FALSE)  # or enum("msoChart", unlist=FALSE)
#' # $`msoChart`
#' # [1] 3
#' # attr(,"family")
#' # [1] "msoShapeType"
#' 
#' enum(msoChart, msoTable)  # or enum(c(msoChart, msoTable))
#'                           # or enum(msoChart, family="Mso")
#' # msoChart   msoTable 
#' #        3         19 
#' # attr(,"family")
#' # [1] "msoShapeType" "msoShapeType"
#' 
#' enum(msoChart, family=c("Xl", "Wd"))  # "msoChart" not in these families
#' # [1] NA
#' 
#' enum(NULL, family="XlAxisGroup")  # return all the enums of "XlAxisGroup"
#' #   xlPrimary xlSecondary 
#' #           1           2 
#' # attr(,"family")
#' # [1] "XlAxisGroup" "XlAxisGroup"
#' 
#' enum(msoChart, bare=TRUE)  # only return bare number
#' # [1] 3
#' 
#' # feel free to use purrr::as_mapper
#' getEnum <- function(family, enum_name) {
#'     purrr:::as_mapper(c(family, enum_name))
#' }
#' getEnum("MsoShapeType", "msoChart")
#' # [1] 3
#' }
enum <- function(enum_name, ..., family=NULL, unlist=TRUE, bare=FALSE){
    if (! exists("ENUM"))
        stop("ENUM data set not found. ",
             "Make sure you have loaded aseshms correctly.")
    
    if (! is.null(family)){
        stopifnot(is.character(family))
        family <- unlist(family)
        match_family <- vapply(family, function(fam){
            grepl(paste0("^", fam), names(ENUM))
        }, logical(length=length(ENUM)))
        match_family <- apply(match_family, 1, any)
        if (! any(match_family)) return(NA)
        ENUM <- ENUM[match_family]
    }
    # unlist ENUM
    x <- unlist(lapply(names(ENUM), function(nm){
        lapply(ENUM[[nm]], structure, family=nm)
    }), recursive=FALSE)
    
    enum_name <- as.character(substitute(enum_name))
    enum_name <- enum_name[!is.na(enum_name)]
    if (enum_name[1] %in% c("c", "list")){
        enum_name <- enum_name[-1]
        dots <- NULL
    }else{
        dots <- as.character(substitute(list(...)))[-1]
    }
    enum_name <- enum_name[! enum_name %in% c("NULL", "NaN")]
    
    if (length(enum_name) > 0){
        enum_name <- c(enum_name, dots)
        if (! any(enum_name %in% names(x)))
            return(NA)
        x <- x[intersect(enum_name, names(x))]
    }
    if (unlist){
        x <- if (bare){
            unname(unlist(x))
        }else{
            structure(unlist(x), family=unname(
                unlist(vapply(x, attr, FUN.VALUE=character(length=1), 
                              which="family"))))
        }
    } else {
        if (bare) x <- as.list(unname(unlist(x)))
    }
    return(x)
}


