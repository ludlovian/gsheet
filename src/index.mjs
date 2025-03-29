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
  const sheets = await getSheetApi()

  debug('Reading %s from %s', range, spreadsheetId)

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER',
    majorDimension: 'ROWS'
  })

  // defensive
  /* c8 ignore start */
  if (response.status !== 200) {
    throw Object.assign(new Error('Failed to read sheet'), { response })
  }
  const data = response.data.values ?? []
  /* c8 ignore stop */
  debug('read %d rows', data.length)
  return data
}

//
// writeSheet
//

export async function writeSheet (spreadsheetId, range, data) {
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

// -----------------------------------------------------------------
//
// clearSheet (spreadsheetId, range)
//

export async function clearSheet (spreadsheetId, range) {
  const sheets = await getSheetApi()

  debug('Clearing %s from %s', range, spreadsheetId)

  const response = await sheets.spreadsheets.values.clear({
    spreadsheetId,
    range
  })

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
