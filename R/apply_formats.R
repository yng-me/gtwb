apply_formats <- function(.data, .formats, ...) {

  for(i in seq_along(.formats)) {

    format <- .formats[[i]]
    cols <- format$cols

    for(j in seq_along(cols)) {

      col <- cols[j]

      .data <- .data |>
        dplyr::mutate(
          !!as.name(col) := dplyr::if_else(
            as.integer(`__row_number__`) %in% format$rows & !is.na(!!as.name(col)),
            format$func$default(as.numeric(!!as.name(col))),
            as.character(!!as.name(col)),
            NA_character_
          )
        )
    }

  }

  return(.data)
}
