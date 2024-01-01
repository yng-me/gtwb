write_footnotes <- function(wb, .footnotes, .start_row, .start_col, ...) {

  restart_at <- .start_row

  if(!is.null(.footnotes)) {

    restart_at <- .start_row + 1

    for(i in seq_along(.footnotes$locname)) {

      note <- .footnotes$footnotes[[i]]

      if("from_markdown" %in% class(note)) {
        note <- note[1]
      }

      wb |> openxlsx::writeData(
        ...,
        x = note,
        startRow = restart_at,
        startCol = .start_col
      )

      restart_at <- restart_at + 1
    }
  }

  return(restart_at)

}


write_source_notes <- function(wb, .source_notes, .start_row, .start_col, ...) {

  restart_at <- .start_row

  if(!is.null(.source_notes)) {

    for(i in seq_along(.source_notes)) {
      note <- .source_notes[[i]]

      if("from_markdown" %in% class(note)) {
        note <- note[1]
      }

      wb |> openxlsx::writeData(
        ...,
        x = note,
        startRow = restart_at,
        startCol = .start_col
      )

      restart_at <- restart_at + 1
    }
  }

  return(restart_at)

}
