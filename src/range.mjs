export class Range {
  // sheet for the range, or possibly undefined if the first sheet
  #sheet

  // top left of the range, might be undefined if the whole sheet
  // Indexes are 1-based
  #top
  #left

  // bottom right of the range (inclusive). Could be undefined if the
  // whole sheet, or a single cell range. Bottom could also be infinity
  // Indexes are 1-based
  #bottom
  #right

  // Creation from strings, values etc

  static fromProps ({ top, left, bottom, right, sheet }) {
    return Object.assign(new Range(), { top, left, bottom, right, sheet })
  }

  static fromRange (input) {
    input = input.split('!')
    const sheet = input.length > 1 ? input.shift() : undefined

    input = input[0].split(':')
    let cell = this.#parseCell(input.shift())
    const left = cell.col
    const top = cell.row

    cell = input.length ? this.#parseCell(input.shift()) : undefined
    const right = cell?.col
    const bottom = cell?.row

    return this.fromProps({ top, left, bottom, right, sheet })
  }

  // Get / Set main properties

  get top () {
    return this.#top
  }

  set top (n) {
    this.#top = n
  }

  get left () {
    return this.#left
  }

  set left (n) {
    this.#left = n
  }

  get bottom () {
    return this.#bottom
  }

  set bottom (n) {
    this.#bottom = n
  }

  get right () {
    return this.#right
  }

  set right (n) {
    this.#right = n
  }

  get sheet () {
    return this.#sheet
  }

  set sheet (s) {
    this.#sheet = s
  }

  get props () {
    return {
      top: this.top,
      left: this.left,
      bottom: this.bottom,
      right: this.right,
      sheet: this.sheet
    }
  }

  // String representation
  //
  toString () {
    return [
      this.sheet,
      [
        [
          !this.left ? '' : Range.b26Encode(this.left),
          !this.top || this.top === Infinity ? '' : this.top
        ].join(''),
        [
          !this.right ? '' : Range.b26Encode(this.right),
          !this.bottom || this.bottom === Infinity ? '' : this.bottom
        ].join('')
      ]
        .filter(Boolean)
        .join(':')
    ]
      .filter(Boolean)
      .join('!')
  }

  // Status
  get hasSheet () {
    return !!this.sheet
  }

  get hasAddress () {
    return !!this.top
  }

  get isRange () {
    return !!this.bottom
  }

  // Derived properties

  get width () {
    return this.right == null ? 1 : this.right - this.left + 1
  }

  set width (n) {
    if (n === undefined) {
      this.bottom = this.right = undefined
    } else {
      this.right = this.left + n - 1
    }
  }

  get height () {
    return this.bottom == null ? 1 : this.bottom - this.top + 1
  }

  set height (n) {
    if (n === undefined) {
      this.bottom = this.right = undefined
    } else {
      this.bottom = this.top + n - 1
    }
  }

  static #parseCell (addr) {
    const match = /^([A-Z]+)(\d+)?$/.exec(addr.toUpperCase())
    if (!match) return {}
    return {
      col: this.b26Decode(match[1]),
      row: match[2] == null ? Infinity : +match[2]
    }
  }

  //
  // The bijective base-26 encoding & decoding
  //
  static b26Encode (col) {
    const toChar = n => String.fromCharCode(64 + n)
    let s = ''
    for (let n = col; n > 0; n = Math.floor(--n / 26)) {
      s = toChar(n % 26 || 26) + s
    }
    return s
  }

  static b26Decode (colName) {
    const codeA = 'A'.codePointAt(0)
    const toDecimal = x => x.codePointAt(0) - codeA + 1
    let m = 1
    let n = 0
    Array.from(colName.toUpperCase())
      .reverse()
      .forEach((s, i) => {
        n += toDecimal(s) * m
        m *= 26
      })
    return n
  }
}

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
