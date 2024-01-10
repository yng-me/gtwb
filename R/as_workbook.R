#' Convert to Excel workbook object
#'
#' @param .data A data frame or tibble.
#' @param .sheet_name Which worksheet to write to. It can be the worksheet index or name. Default is \code{NULL}.
#' @param .start_col A counting number specifying the starting column to write to.
#' @param .start_row A counting number specifying the starting row to write to.
#'
#' @return Workbook object with \code{gt} options applied.
#' @export
#'
#' @examples
#'
#' dplyr::starwars |>
#'   gtx() |>
#'   as_workbook()

as_workbook <- function(
  .data,
  .sheet_name = NULL,
  .start_col = 2,
  .start_row = 2
) {

  if(!inherits(.data, "gt_tbl") & !inherits(.data, "gtx_tbl")) {
    stop("Not a valid gt or gtx object.")
  }

  gtx_facade <- list()
  if(is.null(.sheet_name)) .sheet_name <- "Sheet1"
  wb <- create_workbook(.data, .sheet_name)

  boxhead <- .data[["_boxhead"]] |>
    dplyr::filter(type %in% c("stub", "default")) |>
    dplyr::arrange(dplyr::desc(type)) |>
    dplyr::pull(var)

  from_heading <- wb |>
    write_table_heading(
      .data,
      .start_col = .start_col,
      .start_row = .start_row,
      .col_end = length(boxhead),
      .facade = gtx_facade,
      sheet = .sheet_name
    )

  restart_at <- from_heading$restart_at
  gtx_facade <- from_heading$facade

  gtx_facade <- wb |>
    write_spanners(
      .data,
      .boxhead = boxhead,
      .start_col = .start_col,
      .start_row = restart_at,
      .facade = gtx_facade,
      sheet = .sheet_name
    )

  df_headers <- .data[["_boxhead"]] |>
    dplyr::filter(type == "default") |>
    tidyr::unnest(column_label) |>
    dplyr::pull(column_label) |> t()

  with_stub <- 0
  with_stub_nrow <- .data[["_boxhead"]] |>
    dplyr::filter(type == "stub") |> nrow()

  if(with_stub_nrow > 0) {
    with_stub <- 1
  }

  spanner_level <- 0
  if(nrow(.data[["_spanners"]]) > 0) {
    spanner_level <- max(.data[["_spanners"]]$spanner_level)
  }

  wb |> openxlsx::writeData(
    sheet = .sheet_name,
    x = df_headers,
    startCol = .start_col + with_stub,
    startRow = restart_at + spanner_level,
    colNames = FALSE
  )

  wb |> openxlsx::setRowHeights(
    sheet = .sheet_name,
    rows = restart_at:(restart_at + spanner_level),
    heights = set_row_height(get_value(.data, "column_labels_padding"))
  )

  restart_at <- restart_at + spanner_level
  col_range <- .start_col:(length(boxhead) + 1)

  gtx_facade <- gtx_facade |>
    add_facade(
      style = openxlsx::createStyle(
        valign = "center",
        fontSize = pct_to_pt(
          .px = get_value(.data, "table_font_size"),
          .pct = get_value(.data, "column_labels_font_size")
        )
      ),
      rows = restart_at,
      cols = col_range,
      sheet = .sheet_name
    )

  data_row_start <- restart_at + 1
  row_groups <- .data[["_row_groups"]]

  rows_with_height <- NULL

  if(length(row_groups) > 0) {

    row_group <- .data[["_boxhead"]] |>
      dplyr::filter(type == "row_group") |>
      dplyr::pull(var)

    data_row_start_p <- data_row_start

    row_df_all <- .data[["_data"]] |>
      tibble::rownames_to_column(var = "__row_number__") |>
      apply_formats(.data[["_formats"]])

    for(i in seq_along(row_groups)) {

      row_df <- row_df_all |>
        apply_transforms(
          .data[["_transforms"]],
          .boxhead = boxhead,
          .start_row = data_row_start_p,
          .start_col = .start_col
        ) |>
        apply_col_merge(
          .data[["_col_merge"]],
          .boxhead = boxhead,
          .start_row = restart_at + 1,
          .start_col = .start_col
        ) |>
        dplyr::filter(!!as.name(row_group) == row_groups[i]) |>
        dplyr::select(dplyr::any_of(boxhead)) |>
        dplyr::select(-dplyr::any_of("__row_number__"))

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

      gtx_facade <- gtx_facade |>
        add_border_row_group(
          .data,
          rows = restart_at + 1,
          cols = .start_col:(ncol(row_df) + 1),
          sheet = .sheet_name
        )

      wb |> openxlsx::setRowHeights(
        sheet = .sheet_name,
        rows = restart_at + 1,
        heights = set_row_height(get_value(.data, "row_group_padding"))
      )

      rows_attr <- attributes(row_df)$row_heights
      if(!is.null(rows_attr)) {
        rows_with_height <- c(rows_with_height, rows_attr[1:nrow(row_df)])
      }

      wb |> openxlsx::writeData(
        sheet = .sheet_name,
        x = row_df,
        startCol = .start_col,
        startRow = restart_at + 2,
        colNames = FALSE
      )

      gtx_facade <- gtx_facade |>
        add_facade(
          style = openxlsx::createStyle(
            valign = "center",
            border = "top",
            borderColour = get_value(.data, "table_body_border_bottom_color"),
            wrapText = TRUE
          ),
          rows = (restart_at + 2):(restart_at + nrow(row_df) + 1),
          cols = col_range,
          sheet = .sheet_name
        )

      rows <- row_df_all |>
        dplyr::filter(!!as.name(row_group) == row_groups[i]) |>
        dplyr::pull(`__row_number__`)

      gtx_facade <- apply_styles(
        .data,
        .boxhead = boxhead,
        .rows = rows,
        .start_col = .start_col,
        .start_row = data_row_start_p,
        .facade = gtx_facade,
        sheet = .sheet_name
      )

      restart_at <- restart_at + nrow(row_df) + 1
      data_row_start_p <- data_row_start_p + 1

    }

  } else {

    df <- .data[["_data"]] |>
      tibble::rownames_to_column(var = "__row_number__") |>
      apply_formats(.data[["_formats"]]) |>
      apply_transforms(
        .data[["_transforms"]],
        .boxhead = boxhead,
        .start_row = restart_at + 1,
        .start_col = .start_col
      ) |>
      apply_col_merge(
        .data[["_col_merge"]],
        .boxhead = boxhead,
        .start_row = restart_at + 1,
        .start_col = .start_col
      ) |>
      dplyr::select(dplyr::any_of(boxhead)) |>
      dplyr::select(-dplyr::any_of("__row_number__"))

    wb |> openxlsx::writeData(
      sheet = .sheet_name,
      x = df,
      startCol = .start_col,
      startRow = restart_at + 1,
      colNames = FALSE
    )

    restart_at <- restart_at + nrow(df)

    gtx_facade <- gtx_facade |>
      add_facade(
        style = openxlsx::createStyle(
          valign = "center",
          border = "top",
          borderColour = get_value(.data, "table_body_border_bottom_color"),
          wrapText = TRUE
        ),
        rows = data_row_start:restart_at,
        cols = col_range,
        sheet = .sheet_name
      )
  }

  stub_cols_df <- .data[["_stub_df"]] |>
    dplyr::mutate(start_row = rownum_i + data_row_start)

  if(with_stub > 0) {

    stub_cols_df <- .data[["_stub_df"]] |>
      dplyr::group_by(group_id) |>
      tidyr::nest() |>
      tibble::rownames_to_column(var = "start_row") |>
      dplyr::mutate(start_row = as.integer(start_row) - 1) |>
      tidyr::unnest(data) |>
      dplyr::ungroup() |>
      dplyr::mutate(start_row = start_row + rownum_i + data_row_start) |>
      dplyr::select(dplyr::any_of(names(.data[["_stub_df"]])), start_row)

    stub_cols <- .data[["_boxhead"]] |>
      dplyr::filter(type == "stub") |>
      dplyr::pull(var)

    stub_col <- which(stub_cols %in% boxhead) + .start_col - 1

    for(i in seq_along(stub_cols_df$rownum_i)) {

      indent <- stub_cols_df$indent[i]
      row_stub <- stub_cols_df$start_row[i]

      gtx_facade <- gtx_facade |>
        add_facade(
          style = openxlsx::createStyle(
            border = "right",
            borderColour = get_value(.data, "stub_border_color"),
            borderStyle = set_border_style(get_value(.data, "stub_border_width")),
            indent = as.integer(round(as.integer(indent) / 2.5))
          ),
          rows = row_stub,
          cols = stub_col,
          sheet = .sheet_name
        )
    }
  }

  if(nrow(.data[["_footnotes"]]) > 0) {
    gtx_facade <- gtx_facade |>
      add_facade(
        style = openxlsx::createStyle(
          border = "bottom",
          borderColour = get_value(.data, "table_body_border_bottom_color"),
          borderStyle = set_border_style(get_value(.data, "table_body_border_bottom_width")),
          wrapText = TRUE
        ),
        rows = restart_at,
        cols = col_range,
        sheet = .sheet_name
      )
  }

  wb |> openxlsx::setColWidths(
    sheet = .sheet_name,
    widths = get_value(.data, "table_width"),
    cols = (.start_col + 1):(length(boxhead) + 1)
  )

  for(i in seq_along(boxhead)) {

    col_selected <- .data[["_boxhead"]] |>
      dplyr::filter(var == boxhead[i])

    col_align <- col_selected |> dplyr::pull(column_align)

    gtx_facade <- gtx_facade |>
      add_facade(
        style = openxlsx::createStyle(halign = col_align[1]),
        rows = (data_row_start - 1):restart_at,
        cols = i + .start_col - 1,
        sheet = .sheet_name
      )

    col_width <- col_selected |> dplyr::pull(column_width)

    cols_width_value <- pct_to_pt(
      .px = get_value(.data, "table_font_size"),
      .pct = unlist(col_width),
      .factor = 1/7
    )

    if(is.null(cols_width_value)) {
      cols_width_value <- "auto"
    }

    wb |> openxlsx::setColWidths(
      sheet = .sheet_name,
      widths = cols_width_value,
      cols = i + .start_col - 1
    )
  }

  from_footnotes <- wb |>
    write_end_notes(
      .data,
      .key = "_footnotes",
      .start_col = .start_col,
      .start_row = restart_at,
      .facade = gtx_facade,
      sheet = .sheet_name
    )

  from_source_notes <- wb |>
    write_end_notes(
      .data,
      .key = "_source_notes",
      .start_col = .start_col,
      .start_row = from_footnotes$restart_at,
      .facade = from_footnotes$facade,
      sheet = .sheet_name
    )

  restart_at <- from_source_notes$restart_at

  gtx_facade <- from_source_notes$facade |>
    add_border_outer(
      .data = .data,
      .sheet_name = .sheet_name,
      .cols = col_range,
      .top = .start_row,
      .bottom = restart_at - 1
    )

  if(.start_col > 1) {
    wb |> openxlsx::setColWidths(
      sheet = .sheet_name,
      widths = 2,
      cols = 1
    )
  }

  rows_height <- stub_cols_df$start_row
  rows_height_factor <- 1
  if(!is.null(rows_with_height)) {
    rows_height <- rows_with_height
    rows_height_factor <- 1.25
  }

  wb |> openxlsx::setRowHeights(
    sheet = .sheet_name,
    rows = rows_height,
    heights = set_row_height(
      get_value(.data, "data_row_padding"),
      .factor = rows_height_factor
    )
  )

  wb |> apply_facade(gtx_facade)
  wb |> update_shared_strings(gtx_facade)

  return(wb)
}
