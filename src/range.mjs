export function getColumnName (col) {
  // Convert a column number (1, 2, ..., 26, 27, ...)
  // into a column name (A, B, ..., Z, AA, ...)
  //
  // inspired by bb26
  //
  const toChar = n => String.fromCharCode(64 + n)
  let s = ''
  for (let n = col; n > 0; n = Math.floor(--n / 26)) {
    s = toChar(n % 26 || 26) + s
  }
  return s
}

export function getCellAddress (row, col) {
  return getColumnName(col) + (row === Infinity ? '' : row.toString())
}

export function getRangeAddress (top, left, height, width) {
  const right = left + width - 1
  const bottom = top + height - 1
  return `${getCellAddress(top, left)}:` + `${getCellAddress(bottom, right)}`
}
