import assert from 'node:assert'
import { gzipSync } from 'node:zlib'
import jeeves from '@ludlovian/jeeves'
import Debug from '@ludlovian/debug'

import { getAccessToken } from './auth.js'
import { Range } from './range.js'

const debug = Debug('gsheet:sheets')

const RW_SCOPE = 'https://www.googleapis.com/auth/spreadsheets'
const RO_SCOPE = 'https://www.googleapis.com/auth/spreadsheets.readonly'

const GZIP_MIN = 1024

//  ------------------------------------------------------------------------
//
//  read
//

export async function readSheet (spreadsheetId, ranges, opts = {}) {
  if (Array.isArray(ranges)) return batchRead(spreadsheetId, ranges, opts)
  const data = await batchRead(spreadsheetId, [ranges], opts)
  /* c8 ignore next */
  return data[0] ?? []
}

async function batchRead (spreadsheetId, ranges, opts = {}) {
  ranges = ranges.map(range => Range.from(range).toString())
  const token = await getAccessToken(RO_SCOPE, opts)
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
    authorization: `Bearer ${token}`,
    accept: 'application/json'
  }

  const resp = await jeeves.get(url, { ...opts, headers })
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

export async function writeSheet (spreadsheetId, ranges, datas, opts = {}) {
  if (Array.isArray(ranges)) {
    return batchWrite(spreadsheetId, ranges, datas, opts)
  }
  await batchWrite(spreadsheetId, [ranges], [datas], opts)
}

async function batchWrite (spreadsheetId, ranges, datas, opts = {}) {
  assert(ranges.length === datas.length, 'Mismatch of datas and ranges')
  ranges = ranges.map(r => Range.from(r).toString())

  const token = await getAccessToken(RW_SCOPE, opts)

  debug('updating %s of %s', ranges.join(','), spreadsheetId)

  const payload = {
    valueInputOption: 'RAW',
    data: datas.map((data, ix) => ({
      range: ranges[ix],
      majorDimension: 'ROWS',
      values: data
    }))
  }
  const headers = {
    authorization: `Bearer ${token}`,
    'content-type': 'application/json'
  }

  const body = makeBody(payload, headers)

  const url =
    'https://sheets.googleapis.com' +
    `/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`

  await jeeves.post(url, { ...opts, headers, body }).then(res => res.resume())
}

//  ------------------------------------------------------------------------
//
//  clear
//

export async function clearSheet (spreadsheetId, ranges, opts = {}) {
  if (Array.isArray(ranges)) return batchClear(spreadsheetId, ranges, opts)
  await batchClear(spreadsheetId, [ranges], opts)
}

async function batchClear (spreadsheetId, ranges, opts = {}) {
  ranges = ranges.map(r => Range.from(r).toString())

  const token = await getAccessToken(RW_SCOPE, opts)

  debug('clearing %s of %s', ranges.join(','), spreadsheetId)

  const url =
    'https://sheets.googleapis.com' +
    `/v4/spreadsheets/${spreadsheetId}/values:batchClear`
  const headers = {
    authorization: `Bearer ${token}`,
    'content-type': 'application/json'
  }
  const body = makeBody({ ranges }, headers)

  await jeeves.post(url, { ...opts, headers, body }).then(res => res.resume())
}

//  ------------------------------------------------------------------------
//
//  utilities
//

function makeBody (payload, headers) {
  const str = JSON.stringify(payload)
  if (str.length < GZIP_MIN) {
    const buff = Buffer.from(str)
    headers['content-length'] = buff.length
    return buff
  } else {
    const buff = gzipSync(str)
    headers['content-encoding'] = 'gzip'
    headers['content-length'] = buff.length
    debug('gzip %d to %d', str.length, buff.length)
    return buff
  }
}
