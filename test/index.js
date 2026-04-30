import test from 'node:test'
import assert from 'node:assert/strict'

import {
  readSheet,
  writeSheet,
  clearSheet,
  getLastModified,
  getColumnName,
  getCellAddress,
  getRangeAddress,
  Range,
  convertJsDateToSheets,
  convertSheetsDateToJs
} from '@ludlovian/gsheet'

const spreadsheetId = '1Wka3SAFnCzy_IqipQ-kt9BZwJqAE9TsYxHHK0UsiVkc'

test('API checks', () => {
  test('readSheet', async t => {
    const range = 'Sheet1!A1:B2'
    const act = await readSheet(spreadsheetId, range)
    const exp = [
      ['Foo', 123],
      ['Bar', 45487.75]
    ]
    assert.deepEqual(act, exp)
  })

  test('readSheet - batch', async t => {
    const range = ['Sheet1!A1:B1', 'Sheet1!A2:B2']
    const act = await readSheet(spreadsheetId, range)
    const exp = [[['Foo', 123]], [['Bar', 45487.75]]]
    assert.deepEqual(act, exp)
  })

  test('writeSheet', async t => {
    const range = 'Sheet1!A3:B3'
    const data = [['biz', 'baz']]
    await writeSheet(spreadsheetId, range, data)
  })

  test('writeSheet - batch', async t => {
    const ranges = ['Sheet1!A3:B3', 'Sheet1!A4:B5']
    const datas = [
      [['biz', 'baz']],
      [
        ['piz', 'paz'],
        [7, 17]
      ]
    ]
    await writeSheet(spreadsheetId, ranges, datas)
  })

  test('clearSheet', async t => {
    const range = 'Sheet1!A6:B6'
    await clearSheet(spreadsheetId, range)
  })

  test('clearSheet - batch', async t => {
    const ranges = ['Sheet1!A7', 'Sheet1!B7']
    await clearSheet(spreadsheetId, ranges)
  })

  test('getLastModified', async t => {
    const dt = await getLastModified(spreadsheetId)
    assert(dt instanceof Date)
  })
})

test('Range calcs', () => {
  test('getColumnName', () => {
    const fn = getColumnName
    assert.strictEqual(fn(1), 'A')
    assert.strictEqual(fn(2), 'B')
    assert.strictEqual(fn(26), 'Z')
    assert.strictEqual(fn(27), 'AA')
    assert.strictEqual(fn(52), 'AZ')
    assert.strictEqual(fn(53), 'BA')
  })

  test('getCellAddress', () => {
    const fn = getCellAddress
    assert.strictEqual(fn(1, 1), 'A1')
    assert.strictEqual(fn(3, 4), 'D3')
    assert.strictEqual(fn(Infinity, 2), 'B')
  })

  test('getRangeAddress', () => {
    const fn = getRangeAddress
    assert.strictEqual(fn(1, 2, 3, 4), 'B1:E3')
  })

  test('Range', () => {
    let r
    r = new Range()
    assert.equal(r.toString(), '')
    r = Range.from(r)
    r = Range.from(r.props)
    r = Range.from(r.toString())
    r = Range.from()
    r.sheet = 'Sheet1'
    r.top = 1
    r.left = 1

    assert.equal(r.hasSheet, true)
    assert.equal(r.hasAddress, true)
    assert.equal(r.isRange, false)

    assert.equal(r.toString(), 'Sheet1!A1')
    assert.equal(r.width, 1)
    assert.equal(r.height, 1)

    r.bottom = 1
    r.right = 1

    assert.equal(r.toString(), 'Sheet1!A1:A1')
    assert.equal(r.toUrl(), 'Sheet1!A1%3AA1')

    r.width = 2
    r.height = 2

    assert.equal(r.toString(), 'Sheet1!A1:B2')
    assert.equal(r.width, 2)
    assert.equal(r.height, 2)

    r.width = r.height = undefined
    assert.equal(r.toString(), 'Sheet1!A1')

    r = Range.from('Sheet1!A1:B2')
    const exp = { top: 1, left: 1, bottom: 2, right: 2, sheet: 'Sheet1' }
    assert.deepEqual(r.props, exp)

    r = Range.from('Sheet1!A1:A')
    assert.equal(r.bottom, Infinity)
  })
})

test('Date calcs', () => {
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
/*
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
*/
