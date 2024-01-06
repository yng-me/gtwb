update_shared_strings <- function(wb, ...) {
  from_md <- which(sapply(wb$sharedStrings, \(x) grep("~~~", x)) == 1)

  for(i in seq_along(from_md)) {
    m <- wb$sharedStrings[[from_md[i]]]
    wb$sharedStrings[[from_md[i]]] <- mark_to_sst(m)
  }
}
