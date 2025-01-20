#' Create a *gtx* table object
#'
#' @param .data A data frame or tibble.
#' @param ... Accepts all arguments in \code{as_workbook()}.
#'
#' @return Extended \code{gt} object with additional workbook object configured based on \code{gt} options
#' @export
#'
#' @examples
#'
#' dplyr::starwars |> gtx()
#'

gtx <- function(.data, ...) {

  .data <- gt::gt(.data)
  .data[["_wb"]] <- as_workbook(.data, ...)

  as_gtx_tbl(.data)

}


as_gtx_tbl <- function(.data) {

  class(.data) <- c("gtx_tbl", class(.data))
  .data

}
