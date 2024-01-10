write_table_heading <- function(wb, .data, .start_row, .start_col, .col_end, .facade, ...) {

  col_range <- (.start_col - 1) + (1:.col_end)

  from_title <- wb |>
    add_heading(
      .data = .data,
      .type = "title",
      .start_row = .start_row,
      .start_col = .start_col,
      .col_range = col_range,
      .facade = .facade,
      ...
    )

  from_subtitle <- wb |>
    add_heading(
      .data = .data,
      .type = "subtitle",
      .start_row = from_title$restart_at,
      .start_col = .start_col,
      .col_range = col_range,
      .facade = from_title$facade,
      ...
    )

  row_border_bottom <- from_subtitle$restart_at - 1
  if(is.null(.data[["_heading"]]$title) & is.null(.data[["_heading"]]$subtitle)) {
    row_border_bottom <- row_border_bottom + 1
  }

  .facade <- from_subtitle$facade |>
    add_facade(
      style = openxlsx::createStyle(
        border = "bottom",
        borderColour = get_value(.data, "heading_border_bottom_color"),
        borderStyle = set_border_style(get_value(.data, "heading_border_bottom_width"))
      ),
      rows = row_border_bottom,
      cols = col_range,
      ...
    )

  return(
    list(
      facade = .facade,
      restart_at = from_subtitle$restart_at
    )
  )

}


add_heading <- function(
  wb,
  .data,
  .type,
  .start_row,
  .start_col,
  .col_range,
  .facade,
  ...
) {

  title <- .data[["_heading"]][[.type]]

  if(!is.null(title)) {

    if("from_markdown" %in% class(title)) {
      title <- paste0("[", .start_row, ",", .start_col, "]~~~", title[1])
    }

    wb |> openxlsx::writeData(
      ...,
      x = stringr::str_trim(title),
      startRow = .start_row,
      startCol = .start_col
    )

    wb |> openxlsx::mergeCells(
      ...,
      rows = .start_row,
      cols = .col_range
    )

    wb |> openxlsx::setRowHeights(
      rows = .start_row,
      heights = set_row_height(get_value(.data, "heading_padding")),
      ...
    )

    valign_title <- "center"
    if(!is.null(title) && .type == "title") {
      valign_title <- "bottom"
    }

    if(!is.null(title) && .type == "subtitle") {
      valign_title <- "top"
    }

    fz_var <- paste0("heading_", .type, "_font_size")

    .facade <- .facade |>
      add_facade(
        style = openxlsx::createStyle(
          fontSize = pct_to_pt(
            .px = get_value(.data, "table_font_size"),
            .pct = get_value(.data, fz_var)
          ),
          halign = get_value(.data, "heading_align"),
          valign = valign_title
        ),
        rows = .start_row,
        cols = .start_col,
        ...
      )

    .start_row <- .start_row + 1
  }

  return(
    list(
      facade = .facade,
      restart_at = .start_row
    )
  )

}

