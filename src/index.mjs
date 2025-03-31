import process from 'node:process'
import Debug from '@ludlovian/debug'
import configure from '@ludlovian/configure'

export * from './range.mjs'
export * from './dates.mjs'

const debug = Debug('gsheets')
const scopes = ['https://www.googleapis.com/auth/spreadsheets']

const config = configure('GSHEETS_', {
  credentialsFile: 'creds/credentials.json'
})

// -----------------------------------------------------------------
// External API
//

//
// readSheet (spreadsheetId, range)
//

export async function readSheet (spreadsheetId, range) {
  if (Array.isArray(range)) return batchRead(spreadsheetId, range)

  const sheets = await getSheetApi()

  debug('Reading %s from %s', range, spreadsheetId)

  const params = {
    // path params
    spreadsheetId,
    range,
    // query params
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER',
    majorDimension: 'ROWS'
  }

  const response = await sheets.spreadsheets.values.get(params)
  const body = response.data

  // defensive
  /* c8 ignore start */
  if (response.status !== 200 || !body) {
    throw Object.assign(new Error('Failed to read sheet'), { response })
  }
  const data = body.values ?? []
  debug('read %d rows', data.length)
  /* c8 ignore stop */
  return data
}

async function batchRead (spreadsheetId, ranges) {
  const sheets = await getSheetApi()

  debug('Reading %s from %s', ranges.join(','), spreadsheetId)

  const params = {
    // path params
    spreadsheetId,
    // query params
    ranges,
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER',
    majorDimension: 'ROWS'
  }
  const response = await sheets.spreadsheets.values.batchGet(params)
  const body = response.data

  // defensive
  /* c8 ignore start */
  if (response.status !== 200 || !body) {
    throw Object.assign(new Error('Failed to read sheet'), { response })
  }
  const data = body?.valueRanges.map(vr => vr.values) ?? []
  debug('read %d ranges', data.length)
  /* c8 ignore stop */
  return data
}

//
// writeSheet
//

export async function writeSheet (spreadsheetId, range, data) {
  if (Array.isArray(range)) return batchWrite(spreadsheetId, range, data)

  const sheets = await getSheetApi()

  debug('updating %s of %s with %d rows', range, spreadsheetId, data.length)

  const response = await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: 'RAW',
    requestBody: {
      range,
      majorDimension: 'ROWS',
      values: data
    }
  })
  // defensive
  /* c8 ignore start */
  if (response.status !== 200) {
    throw Object.assign(new Error('Failed to write sheet'), { response })
  }
  /* c8 ignore stop */
}

async function batchWrite (spreadsheetId, ranges, datas) {
  if (ranges.length !== datas.length) {
    throw new Error('Mismatch ranges and data arrays')
  }

  const sheets = await getSheetApi()

  debug('updating %s of %s', ranges.join(','), spreadsheetId)

  const params = {
    // path params
    spreadsheetId,
    // body
    requestBody: {
      valueInputOption: 'RAW',
      data: datas.map((data, ix) => ({
        range: ranges[ix],
        majorDimension: 'ROWS',
        values: data
      }))
    }
  }

  const response = await sheets.spreadsheets.values.batchUpdate(params)
  // defensive
  /* c8 ignore start */
  if (response.status !== 200) {
    throw Object.assign(new Error('Failed to write sheet'), { response })
  }
  /* c8 ignore stop */
}

// -----------------------------------------------------------------
//
// clearSheet (spreadsheetId, range)
//

export async function clearSheet (spreadsheetId, range) {
  if (Array.isArray(range)) return batchClear(spreadsheetId, range)

  const sheets = await getSheetApi()

  debug('Clearing %s from %s', range, spreadsheetId)

  const params = {
    // path params
    spreadsheetId,
    range
  }

  const response = await sheets.spreadsheets.values.clear(params)

  // defensive
  /* c8 ignore start */
  if (response.status !== 200) {
    throw Object.assign(new Error('Failed to read sheet'), { response })
  }
  /* c8 ignore stop */
  return response
}

async function batchClear (spreadsheetId, ranges) {
  const sheets = await getSheetApi()

  debug('Clearing %s from %s', ranges.join(','), spreadsheetId)

  const params = {
    // path params
    spreadsheetId,
    // body
    requestBody: {
      ranges
    }
  }

  const response = await sheets.spreadsheets.values.batchClear(params)

  // defensive
  /* c8 ignore start */
  if (response.status !== 200) {
    throw Object.assign(new Error('Failed to read sheet'), { response })
  }
  /* c8 ignore stop */
  return response
}

// Internal
//
// Load the API

let _sheetApi
async function getSheetApi () {
  if (_sheetApi) return _sheetApi

  const sheetsApi = await import('@googleapis/sheets')
  process.env.GOOGLE_APPLICATION_CREDENTIALS ??= config.credentialsFile
  const auth = new sheetsApi.auth.GoogleAuth({ scopes })
  const authClient = await auth.getClient()
  _sheetApi = sheetsApi.sheets({ version: 'v4', auth: authClient })
  debug('Sheets API loaded')
  return _sheetApi
}
