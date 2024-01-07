add_facade <- function(.facade, ..., .stack = TRUE, sst = list()) {
  if(length(.facade) == 0) .facade <- list()
  .facade[[length(.facade) + 1]] <- rlang::list2(
    ...,
    stack = .stack,
    sst = sst
  )
  .facade
}


apply_facade <- function(wd, .facade) {

  for(i in seq_along(.facade)) {
    fcd <- .facade[[i]]
    wd |>
      openxlsx::addStyle(
        sheet = fcd$sheet,
        style = fcd$style,
        rows = fcd$rows,
        cols = fcd$cols,
        stack = TRUE,
        gridExpand = TRUE
      )
  }
}
