import assert from 'node:assert'
import jeeves from '@ludlovian/jeeves'
import Debug from '@ludlovian/debug'

import { getAccessToken } from './auth.js'
import { Range } from './range.js'

const debug = Debug('gsheets')

const SCOPE = 'https://www.googleapis.com/auth/spreadsheets'

//  ------------------------------------------------------------------------
//
//  read
//

export async function readSheet (spreadsheetId, ranges) {
  if (Array.isArray(ranges)) return batchRead(spreadsheetId, ranges)
  const data = await batchRead(spreadsheetId, [ranges])
  /* c8 ignore next */
  return data[0] ?? []
}

async function batchRead (spreadsheetId, ranges) {
  ranges = ranges.map(range => Range.from(range).toString())
  const token = await getAccessToken(SCOPE)
  debug('Reading %s from %s', ranges.join(','), spreadsheetId)

  const params = new URLSearchParams({
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER',
    majorDimension: 'ROWS'
  })
  ranges.forEach(range => params.append('ranges', range))

  const url =
    'https://sheets.googleapis.com' +
    '/v4/spreadsheets' +
    `/${spreadsheetId}/values:batchGet` +
    `?${params.toString()}`

  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: 'application/json'
  }

  const resp = await jeeves.get(url, { headers })
  const body = await resp.json()

  /* c8 ignore next */
  const data = body?.valueRanges.map(vr => vr.values ?? []) ?? []
  debug('read %d ranges', data.length)
  /* c8 ignore stop */
  return data
}

//  ------------------------------------------------------------------------
//
//  write
//

export async function writeSheet (spreadsheetId, ranges, datas) {
  if (Array.isArray(ranges)) return batchWrite(spreadsheetId, ranges, datas)
  await batchWrite(spreadsheetId, [ranges], [datas])
}

async function batchWrite (spreadsheetId, ranges, datas) {
  assert(ranges.length === datas.length, 'Mismatch of datas and ranges')
  ranges = ranges.map(r => Range.from(r).toString())

  const token = await getAccessToken(SCOPE)

  debug('updating %s of %s', ranges.join(','), spreadsheetId)

  const body = {
    valueInputOption: 'RAW',
    data: datas.map((data, ix) => ({
      range: ranges[ix],
      majorDimension: 'ROWS',
      values: data
    }))
  }

  const url =
    'https://sheets.googleapis.com' +
    `/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`

  const headers = {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json'
  }

  await jeeves.post(url, { headers, body }).then(res => res.text())
}

//  ------------------------------------------------------------------------
//
//  clear
//

export async function clearSheet (spreadsheetId, ranges) {
  if (Array.isArray(ranges)) return batchClear(spreadsheetId, ranges)
  await batchClear(spreadsheetId, [ranges])
}

async function batchClear (spreadsheetId, ranges) {
  ranges = ranges.map(r => Range.from(r).toString())

  const token = await getAccessToken(SCOPE)

  debug('clearing %s of %s', ranges.join(','), spreadsheetId)

  const body = { ranges }
  const url =
    'https://sheets.googleapis.com' +
    `/v4/spreadsheets/${spreadsheetId}/values:batchClear`
  const headers = {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json'
  }

  await jeeves.post(url, { headers, body }).then(res => res.text())
}
