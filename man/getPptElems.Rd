% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/ppt_elem.R
\name{getPptElems}
\alias{getPptElems}
\alias{get_ppt_elems}
\title{Get meta info of the shapes in a Powerpoint document}
\usage{
getPptElems(files = NULL)

get_ppt_elems(files = NULL)
}
\arguments{
\item{files}{character vector of ppt filenames. Default NULL, and a GUI wizard
will launch to help you select the files.}
}
\value{
A list contaning data.frames with columns c("page", "shape", "shapeName",
"typeOfChart", "chartType", "chartIndex).
}
\description{
You can get a data.frame of all the shapes across the ppt, which may help you
program scripts to paint charts.
}
\examples{
\dontrun{
getPptElems(NULL)
}
}
\seealso{
You can refer to \code{ENUM$xlChartType} to see the Enum definitions
of the xlChartTypes.
}
\author{
Yiying Wang, \email{wangy@aetna.com}
}
