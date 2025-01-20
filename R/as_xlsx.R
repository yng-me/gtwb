#' 	Output a *gt* object as Excel
#'
#' @param .data A valid \code{gt} object.
#' @param ... Accepts all arguments in \code{as_workbook()}.
#' @param .filename If \code{NULL} (which is default), will take the title of the table defined in \code{gt::tab_header()} as the filename. If title is not defined, filename will be set to \code{"Book 1.xlsx"}.
#' @param .overwrite If \code{TRUE}, will overwrite any existing file with the same name.
#'
#' @return The same `gt` object. Output as Excel file as a side effect.
#' @export
#'
#' @examples
#' # Source: https://gt.rstudio.com/articles/case-study-gtcars.html
#'
#' # Use dplyr functions to get the car with the best city gas mileage;
#' # this will be used to target the correct cell for a footnote
#'
#' library(gt)
#' library(dplyr)
#'
#' best_gas_mileage_city <- gtcars |>
#'   arrange(desc(mpg_c)) |>
#'   slice(1) |>
#'   mutate(car = paste(mfr, model)) |>
#'   pull(car)
#'
#' # Use dplyr functions to get the car with the highest horsepower
#' # this will be used to target the correct cell for a footnote
#' highest_horsepower <- gtcars |>
#'   arrange(desc(hp)) |>
#'   slice(1) |>
#'   mutate(car = paste(mfr, model)) |>
#'   pull(car)
#'
#' # Define our preferred order for `ctry_origin`
#' order_countries <- c("Germany", "Italy", "United States", "Japan")
#'
#' # Create a display table with `gtcars`, using all of the previous
#' # statements piped together + additional `tab_footnote()` stmts
#' tab <- gtcars |>
#'   arrange(
#'     factor(ctry_origin, levels = order_countries),
#'     mfr, desc(msrp)
#'   ) |>
#'   mutate(car = paste(mfr, model)) |>
#'   select(-mfr, -model) |>
#'   group_by(ctry_origin) |>
#'   gt(rowname_col = "car") |>
#'   cols_hide(columns = c(drivetrain, bdy_style)) |>
#'   cols_move(
#'     columns = c(trsmn, mpg_c, mpg_h),
#'     after = trim
#'   ) |>
#'   tab_spanner(
#'     label = "Performance",
#'     columns = c(mpg_c, mpg_h, hp, hp_rpm, trq, trq_rpm)
#'   ) |>
#'   cols_merge(
#'     columns = c(mpg_c, mpg_h),
#'     pattern = "<<{1}c<br>{2}h>>"
#'   ) |>
#'   cols_merge(
#'     columns = c(hp, hp_rpm),
#'     pattern = "{1}<br>@{2}rpm"
#'   ) |>
#'   cols_merge(
#'     columns = c(trq, trq_rpm),
#'     pattern = "{1}<br>@{2}rpm"
#'   ) |>
#'   cols_label(
#'     mpg_c = "MPG",
#'     hp = "HP",
#'     trq = "Torque",
#'     year = "Year",
#'     trim = "Trim",
#'     trsmn = "Transmission",
#'     msrp = "MSRP"
#'   ) |>
#'   fmt_currency(columns = msrp, decimals = 0) |>
#'   cols_align(
#'     align = "center",
#'     columns = c(mpg_c, hp, trq)
#'   ) |>
#'   tab_style(
#'     style = cell_text(size = px(12)),
#'     locations = cells_body(
#'       columns = c(trim, trsmn, mpg_c, hp, trq)
#'     )
#'   ) |>
#'   text_transform(
#'     locations = cells_body(columns = trsmn),
#'     fn = function(x) {
#'
#'       speed <- substr(x, 1, 1)
#'
#'       type <-
#'         dplyr::case_when(
#'           substr(x, 2, 3) == "am" ~ "Automatic/Manual",
#'           substr(x, 2, 2) == "m" ~ "Manual",
#'           substr(x, 2, 2) == "a" ~ "Automatic",
#'           substr(x, 2, 3) == "dd" ~ "Direct Drive"
#'         )
#'
#'       paste(speed, " Speed<br><em>", type, "</em>")
#'     }
#'   ) |>
#'   tab_header(
#'     title = md("The Cars of **gtcars**"),
#'     subtitle = "These are some fine automobiles"
#'   ) |>
#'   tab_source_note(
#'     source_note = md(
#'       "Source: Various pages within the Edmonds website."
#'     )
#'   ) |>
#'   tab_footnote(
#'     footnote = md("Best gas mileage (city) of all the **gtcars**."),
#'     locations = cells_body(
#'       columns = mpg_c,
#'       rows = best_gas_mileage_city
#'     )
#'   ) |>
#'   tab_footnote(
#'     footnote = md("The highest horsepower of all the **gtcars**."),
#'     locations = cells_body(
#'       columns = hp,
#'       rows = highest_horsepower
#'     )
#'   ) |>
#'   tab_footnote(
#'     footnote = "All prices in U.S. dollars (USD).",
#'     locations = cells_column_labels(columns = msrp)
#'   )
#'
#' # To export `tab` as Excel
#' \dontrun{
#' tab |> as_xlsx()
#' }


as_xlsx <- function(
  .data,
  ...,
  .filename = NULL,
  .overwrite = TRUE
) {

  if(!inherits(.data, "gt_tbl") & !inherits(.data, "gtx_tbl")) {
    .data <- gt(.data, ...)
  }

  if(inherits(.data, "gt_tbl")) {
    wb <- as_workbook(.data, ...)
  } else {
    wb <- .data$wb
  }

  wb |> save_workbook(.data, .filename, overwrite = .overwrite)

  return(.data)

}
