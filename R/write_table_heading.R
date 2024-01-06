write_table_heading <- function(wb, .data, .start_row, .start_col, .col_end, .facade, ...) {

  restart_at <- .start_row
  title <- .data[["_heading"]]$title
  subtitle <- .data[["_heading"]]$subtitle
  col_range <- (.start_col - 1) + (1:.col_end)

  if(!is.null(title)) {

    if("from_markdown" %in% class(title)) {
      title <- paste0("~~~", title[1])
    }

    restart_at <- restart_at + 1

    wb |> openxlsx::writeData(
      x = stringr::str_trim(title),
      startRow = .start_row,
      startCol = .start_col,
      ...
    )

    wb |> openxlsx::mergeCells(
      rows = .start_row,
      cols = col_range,
      ...
    )

    wb |> openxlsx::setRowHeights(
      rows = .start_row,
      heights = set_row_height(get_value(.data, "heading_padding_horizontal")),
      ...
    )

    valign_title <- "center"
    if(!is.null(subtitle)) valign_title <- "bottom"

    .facade <- .facade |>
      add_facade(
        style = openxlsx::createStyle(
          fontSize = pct_to_pt(
            .px = get_value(.data, "table_font_size"),
            .pct = get_value(.data, "heading_title_font_size")
          ),
          halign = get_value(.data, "heading_align"),
          valign = valign_title
        ),
        rows = .start_row,
        cols = .start_col,
        ...
      )

  }

  if(!is.null(subtitle)) {

    if("from_markdown" %in% class(subtitle)) {
      subtitle <- paste0("~~~", subtitle[1])
    }

    restart_at <- restart_at + 1

    wb |> openxlsx::writeData(
      ...,
      x = stringr::str_trim(subtitle),
      startRow = .start_row + 1,
      startCol = .start_col
    )

    wb |> openxlsx::mergeCells(
      ...,
      rows = .start_row + 1,
      cols = col_range
    )

    wb |> openxlsx::setRowHeights(
      rows = .start_row + 1,
      heights = set_row_height(get_value(.data, "heading_padding")),
      ...
    )

    valign_subtitle <- "center"
    if(!is.null(title)) valign_subtitle <- "top"

    .facade <- .facade |>
      add_facade(
        style = openxlsx::createStyle(
          fontSize = pct_to_pt(
            .px = get_value(.data, "table_font_size"),
            .pct = get_value(.data, "heading_subtitle_font_size")
          ),
          halign = get_value(.data, "heading_align"),
          valign = valign_subtitle
        ),
        rows = .start_row + 1,
        cols = .start_col,
        ...
      )

  }

  .facade <- .facade |>
    add_facade(
      style = openxlsx::createStyle(
        border = "top",
        borderColour = get_value(.data, "heading_border_bottom_color"),
        borderStyle = set_border_style(get_value(.data, "heading_border_bottom_width"))
      ),
      rows = restart_at,
      cols = col_range,
      ...
    )

  return(
    list(
      restart_at = restart_at,
      facade = .facade
    )
  )

}
