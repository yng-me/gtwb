add_style <- function(wb, .data, .boxhead, .rows, .start_col, .start_row, ...) {

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

    fontsize <- style$styles[[1]]$cell_text$size

    if(!is.null(fontsize)) {
      fontsize <- pct_to_pt(
        .px = get_value(.data, "table_font_size"),
        .percent = fontsize
      )
    }

    print(style$rownum + .start_row)

    wb |> openxlsx::addStyle(
      style = openxlsx::createStyle(
        fontSize = fontsize,
        wrapText = TRUE
      ),
      rows = style$rownum + .start_row,
      cols = col_which + .start_col - 1,
      gridExpand = TRUE,
      stack = TRUE,
      ...
    )
  }

}
