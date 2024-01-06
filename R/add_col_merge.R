add_col_merge <- function(.data, .col_merge, .boxhead, ...) {

  if(length(.col_merge) == 0) return(.data)

  for(i in seq_along(.col_merge)) {

    merge <- .col_merge[[i]]
    pattern <- merge$pattern
    pattern_new <- pattern
    vars <- merge$vars

    v <- stringr::str_extract_all(pattern, '\\{.*?\\}')[[1]]
    for(j in seq_along(v)) {
      y <- stringr::str_extract_all(v[j], "\\d+")[[1]]
      var <- vars[as.integer(y)]
      pattern_new <- stringr::str_replace(pattern_new, y, var)
    }

    pattern_new <- pattern_new |>
      stringr::str_replace_all("\\s*<br\\s?\\/?>\\s*", "\n") |>
      stringr::str_remove_all("(>|<)+")

    col_selected <- .boxhead[.boxhead %in% vars]

    if(length(col_selected) > 0) {
      .data <- .data |>
        dplyr::mutate(!!as.name(col_selected[1]) := glue::glue(pattern_new))
    }

  }

  return(.data)

}
