#' Title
#'
#' @param .data
#' @param .sheet_name
#' @param .start_col
#' @param .start_row
#' @param .overwrite
#' @param ...
#'
#' @return
#' @export
#'
#' @examples
#'

as_workbook <- function(
  .data,
  .filename,
  .sheet_name = "Sheet 1",
  .start_col = 2,
  .start_row = 2,
  .overwrite = TRUE,
  ...
) {

  wb <- openxlsx::createWorkbook()
  wb |> openxlsx::addWorksheet(.sheet_name, gridLines = FALSE)

  boxhead <- .data[["_boxhead"]] |>
    dplyr::filter(type %in% c("stub", "default")) |>
    dplyr::arrange(dplyr::desc(type)) |>
    dplyr::pull(var)

  restart_at <- wb |>
    write_table_heading(
      .data[["_heading"]],
      .start_col = .start_col,
      .start_row = .start_row,
      .col_end = length(boxhead),
      sheet = .sheet_name
    )

  row_groups <- .data[["_row_groups"]]

  df_headers <- .data[["_boxhead"]] |>
    dplyr::filter(type == "default") |>
    tidyr::unnest(column_label) |>
    dplyr::pull(column_label) |>
    t()

  with_stub <- 0
  with_stub_nrow <- .data[["_boxhead"]] |>
    dplyr::filter(type == "stub") |>
    nrow()

  if(with_stub_nrow > 0) {
    with_stub <- 1
  }


  spanners <- .data[["_spanners"]]
  for(i in seq_along(spanners$spanner_id)) {

  }

  wb |> openxlsx::writeData(
    sheet = .sheet_name,
    x = df_headers,
    startCol = .start_col + with_stub,
    startRow = restart_at,
    colNames = FALSE
  )

  if(!is.null(row_groups)) {

    row_group <- .data[["_boxhead"]] |>
      dplyr::filter(type == "row_group") |>
      dplyr::pull(var)


    for(i in seq_along(row_groups)) {

      row_df <- .data[["_data"]] |>
        dplyr::filter(!!as.name(row_group) == row_groups[i]) |>
        dplyr::select(dplyr::any_of(boxhead))

      # Row group
      wb |> openxlsx::writeData(
        sheet = .sheet_name,
        x = row_groups[i],
        startCol = .start_col,
        startRow = restart_at + 1
      )

      # Merge row group
      wb |> openxlsx::mergeCells(
        sheet = .sheet_name,
        rows = restart_at + 1,
        cols = .start_col:(ncol(row_df) + 1)
      )


      wb |> openxlsx::writeData(
        sheet = .sheet_name,
        x = row_df,
        startCol = .start_col,
        startRow = restart_at + 2,
        colNames = FALSE
      )

      restart_at <- restart_at + nrow(row_df) + 1
    }

  } else {

    df <- .data[["_data"]] |>
      dplyr::select(dplyr::any_of(boxhead))

    wb |> openxlsx::writeData(
      sheet = .sheet_name,
      x = df,
      startCol = .start_col,
      startRow = restart_at + 1,
      colNames = FALSE
    )

    restart_at <- restart_at + nrow(df) + 1

  }

  restart_at <- wb |>
    write_footnotes(
      .data[["_footnotes"]],
      .start_col = .start_col,
      .start_row = restart_at,
      sheet = .sheet_name
    )

  restart_at <- wb |>
    write_source_notes(
      .data[["_source_notes"]],
      .start_col = .start_col,
      .start_row = restart_at,
      sheet = .sheet_name
    )

  wb |> openxlsx::addStyle(
    sheet = .sheet_name,
    style = openxlsx::createStyle(
      border = "top"
    ),
    rows = .start_row,
    cols = .start_col:(length(boxhead) + 1)
  )

  wb |> openxlsx::addStyle(
    sheet = .sheet_name,
    style = openxlsx::createStyle(
      border = "bottom"
    ),
    rows = restart_at - 1,
    cols = .start_col:(length(boxhead) + 1)
  )

  if(.start_col > 1) {
    wb |> openxlsx::setColWidths(
      sheet = .sheet_name,
      widths = 2,
      cols = 1
    )
  }

  wb |> openxlsx::saveWorkbook(file = .filename, overwrite = .overwrite)

  return(.data)

}
