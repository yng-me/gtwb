set_filename <- function(.filename, .heading) {
  if(is.null(.filename)) {
    title <- .heading$title[1]
    if(!is.null(title)) {
      title <- stringr::str_trim(
        stringr::str_replace_all(
          stringr::str_replace_all(title, "[^[:alnum:]]", " "),
          "\\s+",
          " "
        )
      )
      .filename <- paste0(title, ".xlsx")
    } else {
      .filename <- "Book 1.xlsx"
    }
  }
  return(.filename)
}
