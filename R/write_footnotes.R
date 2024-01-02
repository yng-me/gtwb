write_footnotes <- function(wb, .data, .start_row, .start_col, ...) {

  restart_at <- .start_row
  notes <- .data[["_footnotes"]]

  if(!is.null(notes) & nrow(notes) > 0) {

    restart_at <- .start_row + 1

    for(i in seq_along(notes$locname)) {

      note <- notes$footnotes[[i]]

      if("from_markdown" %in% class(note)) {
        note <- paste0("~~~", note[1])
      }

      wb |> openxlsx::writeData(
        x = stringr::str_trim(note),
        startRow = restart_at,
        startCol = .start_col,
        ...
      )

      wb |> openxlsx::addStyle(
        style = openxlsx::createStyle(
          fontSize = percent_to_pt(
            .px = get_value(.data, "footnotes_font_size"),
            .percent = get_value(.data, "table_font_size")
          ),
          valign = "center"
        ),
        rows = restart_at,
        cols = .start_col,
        gridExpand = TRUE,
        stack = TRUE,
        ...
      )

      restart_at <- restart_at + 1
    }
  }

  wb |> openxlsx::setRowHeights(
    rows = .start_row:restart_at,
    heights = set_row_height(get_value(.data, "footnotes_padding_horizontal")),
    ...
  )

  return(restart_at)

}


write_source_notes <- function(wb, .data, .start_row, .start_col, ...) {

  restart_at <- .start_row
  notes <- .data[["_source_notes"]]

  if(!is.null(notes) & length(notes) > 0) {

    for(i in seq_along(notes)) {
      note <- notes[[i]]

      if("from_markdown" %in% class(note)) {
        note <- paste0("~~~", note[1])
      }

      wb |> openxlsx::writeData(
        x = stringr::str_trim(note),
        startRow = restart_at,
        startCol = .start_col,
        ...
      )

      wb |> openxlsx::addStyle(
        style = openxlsx::createStyle(
          fontSize = percent_to_pt(
            .px = get_value(.data, "source_notes_font_size"),
            .percent = get_value(.data, "table_font_size")
          ),
          valign = "center"
        ),
        rows = restart_at,
        cols = .start_col,
        gridExpand = TRUE,
        stack = TRUE,
        ...
      )

      restart_at <- restart_at + 1
    }
  }

  wb |> openxlsx::setRowHeights(
    rows = .start_row:restart_at,
    heights = set_row_height(get_value(.data, "source_notes_padding_horizontal")),
    ...
  )

  return(restart_at)

}
