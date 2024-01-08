apply_transforms <- function(.data, .transforms, .boxhead, .start_row, .start_col, ...) {

  for(i in seq_along(.transforms)) {

    transform <- .transforms[[i]]
    cols <- transform$resolved$colnames

    for(j in seq_along(cols)) {

      col <- cols[j]
      which_col <- which(.boxhead == col)

      .data <- .data |>
        dplyr::mutate(
          is_in_the_row = as.integer(`__row_number__`) %in% transform$resolved$rows,
          prefix = paste0(
            "[",
            .start_row + as.integer(`__row_number__`),
            ",",
            which_col + .start_col - 1,
            "]~~~"
          ),
          !!as.name(col) := dplyr::if_else(
            is_in_the_row & !is.na(!!as.name(col)),
            paste0(prefix, to_mark(transform$fn(!!as.name(col)))),
            !!as.name(col)
          )
        ) |>
        dplyr::select(-c(prefix, is_in_the_row))
    }

  }

  return(.data)
}


to_mark <- function(.text) {
  .text |>
    stringr::str_squish() |>
    stringr::str_trim() |>
    stringr::str_replace_all("\\s*<\\/?b>\\s*", "**") |>
    stringr::str_replace_all("\\s*<\\/?strong>\\s*", "**") |>
    stringr::str_replace_all("\\s*<\\/?i>\\s*", "*") |>
    stringr::str_replace_all("\\s*<\\/?em>\\s*", "*") |>
    stringr::str_replace_all("\\s*<br\\s?\\/?>\\s*", "\n")
}
