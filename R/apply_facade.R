apply_facade <- function(wd, .facade) {

  for(i in seq_along(.facade)) {
    fcd <- .facade[[i]]
    wd |>
      openxlsx::addStyle(
        sheet = fcd$sheet,
        style = fcd$style,
        rows = fcd$rows,
        cols = fcd$cols,
        stack = TRUE,
        gridExpand = TRUE
      )
  }
}

add_facade <- function(.facade, ..., .stack = TRUE, sst = list()) {
  if(length(.facade) == 0) .facade <- list()
  .facade[[length(.facade) + 1]] <- rlang::list2(
    ...,
    stack = .stack,
    sst = sst
  )
  .facade
}


add_border_outer <- function(
  .facade,
  .data,
  .sheet_name,
  .cols,
  .top,
  .bottom
) {

  .facade <- .facade |>
    add_facade(
      style = openxlsx::createStyle(
        border = "top",
        borderColour = get_value(.data, "table_border_top_color"),
        borderStyle = set_border_style(get_value(.data, "table_border_top_width")),
        wrapText = TRUE
      ),
      rows = .top,
      cols = .cols,
      sheet = .sheet_name
    )

  # Bottom border style
  .facade |>
    add_facade(
      style = openxlsx::createStyle(
        border = "bottom",
        borderColour = get_value(.data, "table_border_bottom_color"),
        borderStyle = set_border_style(get_value(.data, "table_border_bottom_width"))
      ),
      rows = .bottom,
      cols = .cols,
      sheet = .sheet_name
    )
}

add_border_row_group <- function(.facade, .data, ...) {

  .facade <- .facade |>
    add_facade(
      style = openxlsx::createStyle(
        border = "bottom",
        borderColour = get_value(.data, "row_group_border_bottom_color"),
        borderStyle = set_border_style(get_value(.data, "row_group_border_bottom_width")),
        wrapText = TRUE,
        valign = "center"
      ),
      ...
    )

  .facade |>
    add_facade(
      style = openxlsx::createStyle(
        border = "top",
        borderColour = get_value(.data, "row_group_border_top_color"),
        borderStyle = set_border_style(get_value(.data, "row_group_border_top_width")),
        valign = "center",
        fontSize = pct_to_pt(
          .px = get_value(.data, "table_font_size"),
          .pct = get_value(.data, "row_group_font_size")
        )
      ),
      ...
    )
}
