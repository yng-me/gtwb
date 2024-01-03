write_spanners <- function(wb, .data, ...) {

  for(i in seq_along(spanners$spanner_id)) {

    span_cols <- which(boxhead %in% spanners$vars[[1]])

    wb |> openxlsx::writeData(
      sheet = .sheet_name,
      x = spanners$spanner_label[i],
      startCol = span_cols[1] + .start_col - 1,
      startRow = restart_at,
    )

    wb |> openxlsx::mergeCells(
      sheet = .sheet_name,
      rows = restart_at,
      cols = span_cols + .start_col - 1
    )

    wb |> openxlsx::addStyle(
      sheet = .sheet_name,
      style = openxlsx::createStyle(
        border = "bottom",
        borderColour = get_value(.data, "column_labels_border_bottom_color"),
        borderStyle = set_border_style(get_value(.data, "column_labels_border_bottom_width")),
        halign = "center",
        valign = "center",
        fontSize = percent_to_pt(
          .px = get_value(.data, "table_font_size"),
          .percent = get_value(.data, "column_labels_font_size")
        )
      ),
      rows = restart_at,
      cols = span_cols + .start_col - 1,
      gridExpand = TRUE,
      stack = TRUE
    )
  }
}
