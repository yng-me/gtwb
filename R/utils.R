get_value <- function(.data, .key) {
  v <- .data[["_options"]] |>
    dplyr::filter(parameter == .key) |>
    dplyr::pull(value)

  if(length(v) > 0) {
    return(v[[1]])
  } else {
    return(NULL)
  }
}

px_to_pt <- function(.px) {
  as.integer(stringr::str_extract(.px, "\\d+")) * 0.75
}


get_font_family <- function(.data) {

  fonts <- get_value(.data, "table_font_names")
  fonts <- fonts[!grepl("emoji|symbol|sans\\-serif", fonts, ignore.case = T)]

  font_family <- "Calibri"

  for(i in seq_along(fonts)) {
    font_family <- systemfonts::system_fonts() |>
      dplyr::filter(family == fonts[i])

    if(nrow(font_family) > 0) {
      font_family <- fonts[i]
    }
  }

  return(font_family)

}

percent_to_pt <- function(.px, .percent) {
  (as.integer(stringr::str_extract(.percent, "\\d+")) / 100) * px_to_pt(.px)
}


set_row_height <- function(.px) {
  16 + (px_to_pt(.px) * 2)
}
