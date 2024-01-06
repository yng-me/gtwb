add_transforms <- function(.data, .transforms, ...) {

  if(length(.transforms) == 0) return(.data)

  for(i in seq_along(.transforms)) {

    transform <- .transforms[[i]]
    cols <- transform$resolved$colnames

    .data <- .data |>
      mutate(
        !!as.name(cols) := paste0(
          '~~~',
          to_mark(transform$fn(!!as.name(cols)))
        )
      )
  }
  return(.data)
}


to_mark <- function(.text) {
  .text |>
    stringr::str_trim() |>
    stringr::str_replace_all("\\s*<\\/?b>\\s*", "**") |>
    stringr::str_replace_all("\\s*<\\/?strong>\\s*", "**") |>
    stringr::str_replace_all("\\s*<\\/?i>\\s*", "*") |>
    stringr::str_replace_all("\\s*<\\/?em>\\s*", "*") |>
    stringr::str_replace_all("\\s*<br\\s?\\/?>\\s*", "\n")
}
