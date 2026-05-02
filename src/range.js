class Range {
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
    // Parsing a range
    //
    // - if theres a ':' or '!' then it is easy
    // - if neither, then it could be
    //    - AZ123 - a cell,       if X{1,3}N+
    //    - AZZ   - a column,     if X{1,3}
    //    - sheet - a whole sheet

    const p = {}
    let ix
    ix = input.indexOf('!')
    if (ix >= 0) {
      p.sheet = input.slice(0, ix)
      input = input.slice(ix + 1)
    } else if (!input.includes(':') && !/^[A-Z]{1,3}\d{0,7}$/.test(input)) {
      p.sheet = input
      input = ''
    }

    ix = input.indexOf(':')
    if (ix >= 0) {
      const cell = this.#parseCell(input.slice(ix + 1), Infinity)
      p.right = cell.col
      p.bottom = cell.row
      input = input.slice(0, ix)
    }

    const cell = this.#parseCell(input, undefined)
    p.left = cell.col
    p.top = cell.row
    return this.fromProps(p)
  }

  static from (data) {
    if (data instanceof Range) return data
    if (typeof data === 'string') return Range.fromRange(data)
    if (data && typeof data === 'object') return Range.fromProps(data)
    return new Range()
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

  toUrl () {
    return encodeURIComponent(this.toString())
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

  static #parseCell (addr, def) {
    if (!addr) return {}
    const match = /^([A-Z]+)(\d+)?$/.exec(addr.toUpperCase())
    if (!match) return {}
    return {
      col: this.b26Decode(match[1]),
      row: match[2] == null ? def : +match[2]
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
    if (!colName) return undefined
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

const getColumnName = Range.b26Encode

function getCellAddress (row, col) {
  return getColumnName(col) + (row === Infinity ? '' : row.toString())
}

function getRangeAddress (top, left, height, width) {
  const right = left + width - 1
  const bottom = top + height - 1
  return `${getCellAddress(top, left)}:` + `${getCellAddress(bottom, right)}`
}

export { Range, getColumnName, getCellAddress, getRangeAddress }
