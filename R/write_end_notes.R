write_end_notes <- function(wb, .data, .key, .start_row, .start_col, .facade, ...) {

  restart_at <- .start_row

  end_notes <- .data[[.key]]

  if(.key == "_footnotes") {
    with_end_note <- isTRUE(!is.null(end_notes) & nrow(end_notes) > 0)
    notes <- end_notes$footnotes
  } else if(.key == "_source_notes") {
    with_end_note <- isTRUE(!is.null(end_notes) & length(end_notes) > 0)
    notes <- end_notes
  }

  if(.key == "_footnotes") {
    restart_at <- .start_row + 1
  }

  var <- stringr::str_remove(.key, "^\\_")
  var_fontsize <- paste0(var, "_font_size")
  var_padding_hr <- paste0(var, "_padding")

  for(i in seq_along(notes)) {

    note <- notes[[i]]

    if("from_markdown" %in% class(note)) {
      note <- paste0("[", restart_at, ",", .start_col, "]~~~", note[1])
    }

    wb |> openxlsx::writeData(
      x = stringr::str_trim(note),
      startRow = restart_at,
      startCol = .start_col,
      ...
    )

    fz <- pct_to_pt(
      .px = get_value(.data, "table_font_size"),
      .pct = get_value(.data, var_fontsize)
    )

    .facade <- .facade |>
      add_facade(
        style = openxlsx::createStyle(
          fontSize = fz,
          valign = "center"
        ),
        rows = restart_at,
        cols = .start_col,
        sst = list(fz = fz),
        ...
      )

    restart_at <- restart_at + 1

  }

  wb |> openxlsx::setRowHeights(
    rows = .start_row:restart_at,
    heights = set_row_height(get_value(.data, var_padding_hr)),
    ...
  )

  return(
    list(
      restart_at = restart_at,
      facade = .facade
    )
  )

}
