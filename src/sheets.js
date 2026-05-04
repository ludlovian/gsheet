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

export async function readSheet (spreadsheetId, ranges) {
  if (Array.isArray(ranges)) return batchRead(spreadsheetId, ranges)
  const data = await batchRead(spreadsheetId, [ranges])
  /* c8 ignore next */
  return data[0] ?? []
}

async function batchRead (spreadsheetId, ranges) {
  ranges = ranges.map(range => Range.from(range).toString())
  const token = await getAccessToken(RO_SCOPE)
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

  const token = await getAccessToken(RW_SCOPE)

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

  await jeeves.post(url, { headers, body }).then(res => res.resume())
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

  const token = await getAccessToken(RW_SCOPE)

  debug('clearing %s of %s', ranges.join(','), spreadsheetId)

  const url =
    'https://sheets.googleapis.com' +
    `/v4/spreadsheets/${spreadsheetId}/values:batchClear`
  const headers = {
    authorization: `Bearer ${token}`,
    'content-type': 'application/json'
  }
  const body = makeBody({ ranges }, headers)

  await jeeves.post(url, { headers, body }).then(res => res.resume())
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
