sst_pattern <- "\\[\\d+\\,\\d+\\]~~~"

update_shared_strings <- function(wb, .facade, ...) {

  with_md <- which(sapply(wb$sharedStrings, \(x) grep(sst_pattern, x)) == 1)

  for(i in seq_along(with_md)) {
    x <- wb$sharedStrings[[with_md[i]]]
    wb$sharedStrings[[with_md[i]]] <- mark_to_sst(x)
  }

  with_rpr <- which(sapply(wb$sharedStrings, \(x) grep("<rPr>", x)) == 1)

  fcd_with_sst <- list()
  for(i in seq_along(.facade)) {
    fcd <- .facade[[i]]
    if(length(fcd$sst) > 0) {
      for(j in seq_along(fcd$rows)) {
        for(k in seq_along(fcd$cols)) {
          fcd_name <- paste0("\\[", fcd$rows[j], "\\,", fcd$cols[k], "\\]~~~")
          fcd_with_sst[[fcd_name]] <- fcd$sst
        }
      }
    }
  }

  for(i in seq_along(with_rpr)) {

    sst <- wb$sharedStrings[[with_rpr[i]]]
    sst_match <- which(stringr::str_detect(sst, names(fcd_with_sst)))

    fontsize <- "<rPr>"

    if(length(sst_match) > 0) {
      fcd_sst <- fcd_with_sst[[sst_match[1]]]

      fz <- fcd_sst$fz

      if(!is.null(fz)) {
        fontsize <- paste0('<rPr><sz val="', fz, '"/>')
      }
    }

    wb$sharedStrings[[with_rpr[i]]] <- wb$sharedStrings[[with_rpr[i]]] |>
      stringr::str_replace_all("<rPr>", fontsize) |>
      stringr::str_remove_all(sst_pattern)
  }
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

  text_as_html <- .text |>
    markdown::mark() |>
    stringr::str_trim() |>
    stringr::str_remove_all("(</?p>|\\n*$)") |>
    stringr::str_remove_all('(^<si><t xml\\:space\\=\\"preserve\\">|</t></si>$)')

  if(length(stringr::str_extract_all(text_as_html, "<(strong|em)>")[[1]]) == 0) {
    return(stringr::str_remove_all(.text, sst_pattern))
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

      new_text_format <- c(
        new_text_format,
        paste0('<r><t xml:space="preserve">', v, "</t></r>")
      )
    }

  }

  paste0("<si>", paste0(new_text_format, collapse = ""), "</si>")

}
