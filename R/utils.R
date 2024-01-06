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

  if(grepl("%$", .pct)) {
    (px_to_pt(.px) * (to_int(.pct) / 100)) * .factor
  } else if(grepl("px$", .pct)) {
    px_to_pt(.pct) * .factor
  } else {
    return(NULL)
  }
}

set_row_height <- function(.px) {
  16 + (px_to_pt(.px) * 2)
}


mark_to_sst <- function(.text) {

  get_emp <- function(.emp) {

    v <- stringr::str_extract(.emp, "<(strong|em)>") |>
      purrr::pluck(1) |>
      purrr::discard(~ . == "")

    if(is.na(v)) {
      fmt <- ""
    } else if(v == "<strong>") {
      fmt <- "<b/>"
    } else if(v == "<em>") {
      fmt <- "<i/>"
    } else if(v == "<code>") {
      fmt <- "<fz value=\"Courier New\" />"
    } else {
      fmt <- ""
    }
    return(fmt)
  }

  text_as_html <- stringr::str_remove_all(
    stringr::str_remove_all(
      stringr::str_trim(markdown::mark(.text)), "(~~~|</?p>|\\n*)"
    ),
    '(^<si><t xml\\:space\\=\\"preserve\\">|</t></si>$)'
  )

  if(length(stringr::str_extract_all(text_as_html, "<(strong|em)>")[[1]]) == 0) {
    return(stringr::str_remove_all(.text, "~~~"))
  }

  text_splits <- stringr::str_split(text_as_html, "</(strong|em)>") |>
    purrr::pluck(1) |>
    purrr::discard(~ . == "")

  new_text_format <- NULL

  for(i in seq_along(text_splits)) {

    v <- stringr::str_split(text_splits[i], "<(strong|em)>") |>
      purrr::pluck(1) |>
      purrr::discard(~ . == "")

    if(length(v) == 2) {

      v1 <- paste0('<r><t xml:space="preserve">', v[1], "</t></r>")
      v2 <- paste0(
        '<r><rPr>',
        get_emp(text_splits[i]),
        '</rPr><t xml:space="preserve">',
        v[2],
        "</t></r>"
      )

      new_text_format <- c(new_text_format, paste0(v1, v2))

    } else {

      new_text_format <- c(new_text_format, paste0('<r><t xml:space="preserve">', v, "</t></r>"))
    }

  }

  return(paste0("<si>", paste0(new_text_format, collapse = ""), "</si>"))

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
