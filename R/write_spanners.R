write_spanners <- function(wb, .data, .boxhead, .start_col, .start_row, ...) {

  restart_at <- .start_row
  spanners <- .data[["_spanners"]]
  spanner_levels <- rev(1:max(spanners$spanner_level))

  for(i in seq_along(spanner_levels)) {

    spanners_by_level <- spanners |>
      dplyr::filter(spanner_level == i)


    for(j in seq_along(spanners_by_level$spanner_id)) {

      span_cols <- which(.boxhead %in% spanners_by_level$vars[[1]])

      wb |> openxlsx::writeData(
        x = spanners_by_level$spanner_label[j],
        startCol = span_cols[1] + .start_col - 1,
        startRow = restart_at,
        ...
      )

      wb |> openxlsx::mergeCells(
        rows = restart_at,
        cols = span_cols + .start_col - 1,
        ...
      )

      wb |> openxlsx::addStyle(
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
        stack = TRUE,
        ...
      )

    }

    restart_at <- restart_at + 1
  }

}
