import { suite, test } from 'node:test'
import assert from 'node:assert/strict'

import {
  readSheet,
  writeSheet,
  getColumnName,
  getCellAddress,
  getRangeAddress,
  convertJsDateToSheets,
  convertSheetsDateToJs
} from '@ludlovian/gsheet'

const spreadsheetId = '1Wka3SAFnCzy_IqipQ-kt9BZwJqAE9TsYxHHK0UsiVkc'

suite('Sheet access', { concurrency: false }, () => {
  test('read sheet', async () => {
    const range = 'Sheet1!A1:B2'
    const exp = [
      ['Foo', 123],
      ['Bar', 45487.75]
    ]
    const act = await readSheet(spreadsheetId, range)
    assert.deepStrictEqual(act, exp)
  })

  test('write sheet', async () => {
    const range = 'Sheet1!A4:B5'
    const data = [
      ['Fizz', 123],
      ['Buzz', 456]
    ]
    await writeSheet(spreadsheetId, range, data)
    const act = await readSheet(spreadsheetId, range)
    assert.deepStrictEqual(act, data)

    const blank = [
      ['', ''],
      ['', '']
    ]
    await writeSheet(spreadsheetId, range, blank)
  })
})

suite('Range calcs', () => {
  let fn
  fn = getColumnName
  assert.strictEqual(fn(1), 'A')
  assert.strictEqual(fn(2), 'B')
  assert.strictEqual(fn(26), 'Z')
  assert.strictEqual(fn(27), 'AA')
  assert.strictEqual(fn(52), 'AZ')
  assert.strictEqual(fn(53), 'BA')

  fn = getCellAddress
  assert.strictEqual(fn(1, 1), 'A1')
  assert.strictEqual(fn(3, 4), 'D3')
  assert.strictEqual(fn(Infinity, 2), 'B')

  fn = getRangeAddress
  assert.strictEqual(fn(1, 2, 3, 4), 'B1:E3')
})

suite('Date calcs', () => {
  const summerSerial = 45487.75
  const winterSerial = 45307.25
  const summerDate = new Date(2024, 6, 14, 18) // 2024-07-14 18:00:00
  const winterDate = new Date(2024, 0, 16, 6) //  2024-01-06 06:00:00

  const areClose = (a, b) => Math.abs(a - b) < 1e-9

  assert(areClose(+convertSheetsDateToJs(summerSerial), +summerDate))
  assert(areClose(+convertSheetsDateToJs(winterSerial), +winterDate))
  assert(areClose(convertJsDateToSheets(summerDate), summerSerial))
  assert(areClose(convertJsDateToSheets(winterDate), winterSerial))
})
