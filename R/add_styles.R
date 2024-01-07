add_styles <- function(.data, .boxhead, .rows, .start_col, .start_row, .facade, ...) {

  styles <- .data[["_styles"]] |>
    dplyr::filter(locname == "data")

  if(nrow(styles) == 0) return(invisible())

  colnames <- styles |>
    dplyr::distinct(colname) |>
    pull(colname)

  for(i in seq_along(colnames)) {
    col <- colnames[i]
    col_which <- which(.boxhead == col)

    style <- styles |>
      dplyr::filter(colname == col) |>
      dplyr::filter(rownum %in% as.integer(.rows))

    fz <- style$styles[[1]]$cell_text$size

    sst <- list()
    if(!is.null(fz)) {
      fz <- pct_to_pt(
        .px = get_value(.data, "table_font_size"),
        .pct = fz
      )
      sst <- list(fz = fz)
    }

    .facade <- .facade |>
      add_facade(
        style = openxlsx::createStyle(fontSize = fz, wrapText = TRUE),
        rows = style$rownum + .start_row,
        cols = col_which + .start_col - 1,
        sst = sst,
        ...
      )
  }

  return(.facade)

}
