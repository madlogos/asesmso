% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/get_chart_dims.R
\name{getMsoChartDims}
\alias{getMsoChartDims}
\alias{get_msochart_dims}
\title{Get dims parameters of the matrix yielded from the msoChart object}
\usage{
getMsoChartDims(series)

get_msochart_dims(series)
}
\arguments{
\item{series}{series yielded from msoChart object by \code{getMsoChartSeries}}
}
\value{
A list comprising of 4 vectors: \describe{
\item{NamRows}{Row of Name part} \item{NamCols}{Col of Name part}
\item{ValRows}{Row of Values part} \item{ValCols}{Col of Values part}
\item{LblRows}{Row of XValues part} \item{lblCols}{Col of XValues part}}
}
\description{
Get the indices of rows and columnss involved in msoChart object. The function
will parse the series list yieled by \code{\link{getMsoChartSeries}}, and then
get the \strong{original} parameters out of the matrix. Please note that that
matrix may be different from that should be according to the actual chart.
}
\examples{
\dontrun{
ppt <- RDCOMClient::COMCreate("Powerpoint.Application")
pres <- ppt$Presentations()$Open(<some ppt file>)
slide <- pres$Slides(2)  # the 2nd slide
shape <- slide$Shapes(3)  # the 3rd shape
if (shape[['HasChart']] == -1) {  # if the shape contains msoChart
    chart <- shape[['Chart']]
    series <- getMsoChartSeries(chart)
    getMsoChartDims(series)
}
}
}
\seealso{
\code{\link{getMsoChartSeries}}
}
\author{
Yiying Wang, \email{wangy@aetna.com}
}
