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

to_int <- function(.value) {
  as.integer(stringr::str_extract(.value, "\\d+"))
}

px_to_pt <- function(.px) {
  if(grepl("px$", .px)) {
    to_int(.px) * 0.75
  } else {
    return(NULL)
  }
}


pct_to_pt <- function(.px, .pct, .factor = 1) {

  if(is.null(.pct)) return(NULL)

  if(grepl("%$", .pct)) {
    (px_to_pt(.px) * (to_int(.pct) / 100)) * .factor
  } else if(grepl("px$", .pct)) {
    px_to_pt(.pct) * .factor
  } else {
    return(NULL)
  }
}

set_row_height <- function(.px, .factor = 1) {
  (18 + (px_to_pt(.px) * 2)) * .factor
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

set_border_style <- function(.width) {

  border_width <- to_int(.width)

  if(border_width > 1 & border_width < 4) {
    x <- "medium"
  } else if (border_width >= 4) {
    x <- "thick"
  } else {
    x <- "thin"
  }

  return(x)
}
