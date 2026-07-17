import { getAccessToken } from './auth.js'
import jeeves from '@ludlovian/jeeves'
import Debug from '@ludlovian/debug'

const debug = Debug('gsheet:drive')

const SCOPE = 'https://www.googleapis.com/auth/drive.metadata.readonly'

// -----------------------------------------------------------------
//
// getLastModified (id)
//

export async function getLastModified (fileId, opts = {}) {
  const token = await getAccessToken(SCOPE, opts)

  debug('Getting lastMod from %s', fileId)

  const params = new URLSearchParams({
    fields: 'modifiedTime',
    supportsAllDrives: true
  })
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: 'application/json'
  }
  const url =
    'https://www.googleapis.com' +
    `/drive/v3/files/${fileId}?${params.toString()}`

  const resp = await jeeves.get(url, { ...opts, headers })
  const body = await resp.json()

  /* c8 ignore next */
  if (!body?.modifiedTime) return null
  return new Date(body.modifiedTime)
}
