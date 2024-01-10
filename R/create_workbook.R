create_workbook <- function(.data, .sheet_name = NULL) {

  wb <- openxlsx::createWorkbook()

  wb |> openxlsx::addWorksheet(
    .sheet_name,
    gridLines = FALSE,
    orientation = get_value(.data, "page_orientation")
  )

  wb |> openxlsx::modifyBaseFont(
    fontColour = get_value(.data, "table_font_color"),
    fontSize = px_to_pt(get_value(.data, "table_font_size")),
    fontName = get_font_family(.data)
  )

  return(wb)
}


save_workbook <- function(wb, .data, .filename, ...) {
  wb |> openxlsx::saveWorkbook(
    file = set_filename(.filename, .data[["_heading"]]),
    ...
  )
}
