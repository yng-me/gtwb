#' 	Output a *gt* object as Excel
#'
#' @param .data A valid \code{gt} object.
#' @param .filename If \code{NULL} (which is default), will take the title of the table defined in \code{gt::tab_header()} as the filename. If title is not defined, filename will be set to \code{"Book 1.xlsx"}.
#' @param .sheet_name Which worksheet to write to. It can be the worksheet index or name. Default is \code{"Sheet 1"}.
#' @param .start_col A counting number specifying the starting column to write to.
#' @param .start_row A counting number specifying the starting row to write to.
#' @param .overwrite If \code{TRUE}, will overwrite any existing file with the same name.
#' @param ... For future implementation.
#'
#' @return The same `gt` object. Output as Excel file as a side effect.
#' @export
#'
#' @examples
#' # Source: https://gt.rstudio.com/articles/case-study-gtcars.html
#'
#' # Use dplyr functions to get the car with the best city gas mileage;
#' # this will be used to target the correct cell for a footnote
#'
#' library(gt)
#' library(dplyr)
#'
#' best_gas_mileage_city <- gtcars |>
#'   arrange(desc(mpg_c)) |>
#'   slice(1) |>
#'   mutate(car = paste(mfr, model)) |>
#'   pull(car)
#'
#' # Use dplyr functions to get the car with the highest horsepower
#' # this will be used to target the correct cell for a footnote
#' highest_horsepower <- gtcars |>
#'   arrange(desc(hp)) |>
#'   slice(1) |>
#'   mutate(car = paste(mfr, model)) |>
#'   pull(car)
#'
#' # Define our preferred order for `ctry_origin`
#' order_countries <- c("Germany", "Italy", "United States", "Japan")
#'
#' # Create a display table with `gtcars`, using all of the previous
#' # statements piped together + additional `tab_footnote()` stmts
#' tab <- gtcars |>
#'   arrange(
#'     factor(ctry_origin, levels = order_countries),
#'     mfr, desc(msrp)
#'   ) |>
#'   mutate(car = paste(mfr, model)) |>
#'   select(-mfr, -model) |>
#'   group_by(ctry_origin) |>
#'   gt(rowname_col = "car") |>
#'   cols_hide(columns = c(drivetrain, bdy_style)) |>
#'   cols_move(
#'     columns = c(trsmn, mpg_c, mpg_h),
#'     after = trim
#'   ) |>
#'   tab_spanner(
#'     label = "Performance",
#'     columns = c(mpg_c, mpg_h, hp, hp_rpm, trq, trq_rpm)
#'   ) |>
#'   cols_merge(
#'     columns = c(mpg_c, mpg_h),
#'     pattern = "<<{1}c<br>{2}h>>"
#'   ) |>
#'   cols_merge(
#'     columns = c(hp, hp_rpm),
#'     pattern = "{1}<br>@{2}rpm"
#'   ) |>
#'   cols_merge(
#'     columns = c(trq, trq_rpm),
#'     pattern = "{1}<br>@{2}rpm"
#'   ) |>
#'   cols_label(
#'     mpg_c = "MPG",
#'     hp = "HP",
#'     trq = "Torque",
#'     year = "Year",
#'     trim = "Trim",
#'     trsmn = "Transmission",
#'     msrp = "MSRP"
#'   ) |>
#'   fmt_currency(columns = msrp, decimals = 0) |>
#'   cols_align(
#'     align = "center",
#'     columns = c(mpg_c, hp, trq)
#'   ) |>
#'   tab_style(
#'     style = cell_text(size = px(12)),
#'     locations = cells_body(
#'       columns = c(trim, trsmn, mpg_c, hp, trq)
#'     )
#'   ) |>
#'   text_transform(
#'     locations = cells_body(columns = trsmn),
#'     fn = function(x) {
#'
#'       speed <- substr(x, 1, 1)
#'
#'       type <-
#'         dplyr::case_when(
#'           substr(x, 2, 3) == "am" ~ "Automatic/Manual",
#'           substr(x, 2, 2) == "m" ~ "Manual",
#'           substr(x, 2, 2) == "a" ~ "Automatic",
#'           substr(x, 2, 3) == "dd" ~ "Direct Drive"
#'         )
#'
#'       paste(speed, " Speed<br><em>", type, "</em>")
#'     }
#'   ) |>
#'   tab_header(
#'     title = md("The Cars of **gtcars**"),
#'     subtitle = "These are some fine automobiles"
#'   ) |>
#'   tab_source_note(
#'     source_note = md(
#'       "Source: Various pages within the Edmonds website."
#'     )
#'   ) |>
#'   tab_footnote(
#'     footnote = md("Best gas mileage (city) of all the **gtcars**."),
#'     locations = cells_body(
#'       columns = mpg_c,
#'       rows = best_gas_mileage_city
#'     )
#'   ) |>
#'   tab_footnote(
#'     footnote = md("The highest horsepower of all the **gtcars**."),
#'     locations = cells_body(
#'       columns = hp,
#'       rows = highest_horsepower
#'     )
#'   ) |>
#'   tab_footnote(
#'     footnote = "All prices in U.S. dollars (USD).",
#'     locations = cells_column_labels(columns = msrp)
#'   )
#'
#' # To export `tab` as Excel
#' \dontrun{
#' tab |> as_xlsx()
#' }


as_xlsx <- function(
  .data,
  .filename = NULL,
  .sheet_name = "Sheet 1",
  .start_col = 2,
  .start_row = 2,
  .overwrite = TRUE,
  ...
) {

  if(!("gt_tbl" %in% class(.data))) {
    stop("Not a valid gt object.")
  }

  gtx_facade <- list()
  type <- NULL
  var <- NULL
  column_label <- NULL

  wb <- openxlsx::createWorkbook()

  wb |> modifyBaseFont(
    fontColour = get_value(.data, "table_font_color"),
    fontSize = px_to_pt(get_value(.data, "table_font_size")),
    fontName = get_font_family(.data)
  )

  wb |> openxlsx::addWorksheet(
    .sheet_name,
    gridLines = FALSE,
    orientation = get_value(.data, "page_orientation")
  )

  wb_style <- wb

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

  spanner_level <- max(.data[["_spanners"]]$spanner_level)

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
    heights = set_row_height(get_value(.data, "column_labels_padding_horizontal"))
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

  if(length(row_groups) > 0) {

    row_group <- .data[["_boxhead"]] |>
      dplyr::filter(type == "row_group") |>
      dplyr::pull(var)

    data_row_start_p <- data_row_start

    for(i in seq_along(row_groups)) {

      row_df <- .data[["_data"]] |>
        add_col_merge(.data[["_col_merge"]], boxhead) |>
        add_transforms(.data[["_transforms"]]) |>
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

      gtx_facade <- gtx_facade |>
        add_facade(
          style = openxlsx::createStyle(
            border = "bottom",
            borderColour = get_value(.data, "row_group_border_bottom_color"),
            borderStyle = set_border_style(get_value(.data, "row_group_border_bottom_width"))
          ),
          rows = restart_at + 1,
          cols = .start_col:(ncol(row_df) + 1),
          sheet = .sheet_name
        )

      gtx_facade <- gtx_facade |>
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
          rows = restart_at + 1,
          cols = .start_col:(ncol(row_df) + 1),
          sheet = .sheet_name
        )

      wb |> openxlsx::setRowHeights(
        sheet = .sheet_name,
        rows = restart_at + 1,
        heights = set_row_height(get_value(.data, "row_group_padding_horizontal"))
      )

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

      rows <- .data[["_data"]] |>
        tibble::rownames_to_column() |>
        dplyr::filter(!!as.name(row_group) == row_groups[i]) |>
        pull(rowname)

      gtx_facade <- add_styles(
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
      add_col_merge(.data[["_col_merge"]], boxhead) |>
      add_transforms(.data[["_transforms"]]) |>
      dplyr::select(dplyr::any_of(boxhead))

    wb |> openxlsx::writeData(
      sheet = .sheet_name,
      x = df,
      startCol = .start_col,
      startRow = restart_at + 1,
      colNames = FALSE
    )

    restart_at <- restart_at + nrow(df) + 1

    gtx_facade <- gtx_facade |>
      add_facade(
        style = openxlsx::createStyle(
          valign = "center",
          border = "top",
          borderColour = get_value(.data, "table_body_border_bottom_color"),
          wrapText = TRUE
        ),
        rows = data_row_start:(restart_at - 1),
        cols = col_range,
        sheet = .sheet_name
      )
  }

  # wb |> openxlsx::setRowHeights(
  #   sheet = .sheet_name,
  #   rows = data_row_start:restart_at,
  #   heights = set_row_height(get_value(.data, "data_row_padding_horizontal"))
  # )

  add_row <- 1
  if(nrow(.data[["_footnotes"]]) > 0) {
    add_row <- 0
  }

  if(with_stub > 0) {

    stub_cols <- .data[["_boxhead"]] |>
      dplyr::filter(type == "stub") |>
      pull(var)

    gtx_facade <- gtx_facade |>
      add_facade(
        style = openxlsx::createStyle(
          border = "right",
          borderColour = get_value(.data, "stub_border_color"),
          borderStyle = set_border_style(get_value(.data, "stub_border_width"))
        ),
        rows = data_row_start:(restart_at - add_row),
        cols = which(stub_cols %in% boxhead) + .start_col - 1,
        sheet = .sheet_name
      )
  }

  gtx_facade <- gtx_facade |>
    add_facade(
      style = openxlsx::createStyle(
        border = "bottom",
        borderColour = get_value(.data, "table_body_border_bottom_color"),
        borderStyle = set_border_style(get_value(.data, "table_body_border_bottom_width"))
      ),
      rows = restart_at - add_row,
      cols = col_range,
      sheet = .sheet_name
    )

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

    wb |> setColWidths(
      sheet = .sheet_name,
      widths = pct_to_pt(
        .px = get_value(.data, "table_font_size"),
        .pct = unlist(col_width),
        .factor = 1/8
      ),
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

  # Top border style
  gtx_facade <- from_source_notes$facade |>
    add_facade(
      style = openxlsx::createStyle(
        border = "top",
        borderColour = get_value(.data, "table_border_top_color"),
        borderStyle = set_border_style(get_value(.data, "table_border_top_width")),
        wrapText = TRUE
      ),
      rows = .start_row,
      cols = col_range,
      sheet = .sheet_name
    )

  # Bottom border style
  gtx_facade <- gtx_facade |>
    add_facade(
      style = openxlsx::createStyle(
        border = "bottom",
        borderColour = get_value(.data, "table_border_bottom_color"),
        borderStyle = set_border_style(get_value(.data, "table_border_bottom_width"))
      ),
      rows = restart_at - 1,
      cols = col_range,
      sheet = .sheet_name
    )

  if(.start_col > 1) {
    wb |> openxlsx::setColWidths(
      sheet = .sheet_name,
      widths = 2,
      cols = 1
    )
  }

  wb |> update_shared_strings()
  wb |> apply_facade(gtx_facade)

  assign('gtx_facade', gtx_facade, envir = globalenv())

  wb |> openxlsx::saveWorkbook(
    file = set_filename(.filename, .data[["_heading"]]),
    overwrite = .overwrite
  )

  return(.data)

}


