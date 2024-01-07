add_formats <- function(.data, .formats, ...) {

  if(length(.formats) == 0) return(.data)

  for(i in seq_along(.formats)) {

    format <- .formats[[i]]

    cols <- format$cols

    for(j in seq_along(cols)) {

      col <- cols[j]

      .data <- .data |>
        dplyr::mutate(
          !!as.name(col) := dplyr::if_else(
            as.integer(`__row_number__`) %in% format$rows,
            format$func$default(!!as.name(col)),
            as.character(!!as.name(col))
          )
        )
    }

  }

  return(.data)
}
