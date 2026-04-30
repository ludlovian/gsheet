import assert from 'node:assert'
import { getAccessToken } from './auth.js'
import Debug from '@ludlovian/debug'

import { fetch } from './fetch.js'
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

  const { body } = await fetch({
    hostname: 'sheets.googleapis.com',
    path:
      '/v4/spreadsheets' +
      `/${spreadsheetId}/values:batchGet` +
      `?${params.toString()}`,
    method: 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json'
    }
  })

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

  await fetch({
    hostname: 'sheets.googleapis.com',
    path: `/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`,
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body
  })
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

  await fetch({
    hostname: 'sheets.googleapis.com',
    path: `/v4/spreadsheets/${spreadsheetId}/values:batchClear`,
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body
  })
}
