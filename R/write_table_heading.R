write_table_heading <- function(wb, .heading, .start_row, .start_col, .col_end, ...) {

  restart_at <- .start_row
  title <- .heading$title
  subtitle <- .heading$subtitle
  merge_cols <- (.start_col - 1) + (1:.col_end)

  if(!is.null(title)) {

    if("from_markdown" %in% class(title)) {
      title <- title[1]
    }

    restart_at <- restart_at + 1

    wb |> openxlsx::writeData(
      ...,
      x = title,
      startRow = .start_row,
      startCol = .start_col
    )

    wb |> openxlsx::mergeCells(
      ...,
      rows = .start_row,
      cols = merge_cols
    )

  }

  if(!is.null(subtitle)) {

    if("from_markdown" %in% class(subtitle)) {
      subtitle <- subtitle[1]
    }

    restart_at <- restart_at + 1

    wb |> openxlsx::writeData(
      ...,
      x = subtitle,
      startRow = .start_row + 1,
      startCol = .start_col
    )

    wb |> openxlsx::mergeCells(
      ...,
      rows = .start_row + 1,
      cols = merge_cols
    )

  }

  return(restart_at)

}
