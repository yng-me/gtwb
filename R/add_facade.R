add_facade <- function(wdf, ...) {
  wdf |> openxlsx::addStyle(...)
}

# add_facade <- function(.facade, ...) {
#   if(length(.facade) == 0) .facade <- list()
#   facade_new <- rlang::list2(...)
#   print(names(facade_new))
#   rlang::set_names(facade_new, names(facade_new))
#
#   .facade[[length(.facade) + 1]] <- facade_new
#
#   .facade
# }


apply_facade <- function(wd, .facade) {
  for(i in seq_along(.facade)) {
    fcd <- .facade[[i]]
    wd |> openxlsx::addStyle(
      sheet = fcd$sheet,
      style = fcd$style,
      rows = fcd$rows,
      cols = fcd$cols,
      gridExpand = TRUE,
      stack = TRUE
    )
  }
}
