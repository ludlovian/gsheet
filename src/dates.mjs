const UNIX_EPOCH_IN_XL = 25_569
const MS_PER_DAY = 86_400_000
const MS_PER_MIN = 60_000
const MINS_PER_DAY = 1_440

// Excel <--> JS
//
// Excel / Sheets dates are purely local wall-time.
//
// So local noon, whereever you are in the world will always be 0.5
//
// Julian & JS dates are points-in-time, where 0.5 represents noon GMT
//
// Conversion steps:
//  - start with ms since Unix Epoch for this point in time
//  - convert to number of days since Excel epoch
//  - move the clock-time for local tine
//
export function convertJsDateToSheets (d) {
  const minsOffset = d.getTimezoneOffset()
  return +d / MS_PER_DAY + UNIX_EPOCH_IN_XL - minsOffset / MINS_PER_DAY
}

// From serial to JS we reverse it:
// - start with days since XL epoch
// - convert to days since UX epoch
// - convert to ms since UX epoch (but still in local time)
// - make a date, get the TZ offset, and adjust

export function convertSheetsDateToJs (serial) {
  const ms = (serial - UNIX_EPOCH_IN_XL) * MS_PER_DAY
  const d = new Date(ms)
  return new Date(+d + d.getTimezoneOffset() * MS_PER_MIN)
}
