import { getAccessToken } from './auth.js'
import Debug from '@ludlovian/debug'

import { fetch } from './fetch.js'

const debug = Debug('gsheets')

const SCOPE = 'https://www.googleapis.com/auth/drive.metadata.readonly'

// -----------------------------------------------------------------
//
// getLastModified (id)
//

export async function getLastModified (fileId) {
  const token = await getAccessToken(SCOPE)

  debug('Getting lastMod from %s', fileId)

  const params = new URLSearchParams({
    fields: 'modifiedTime',
    supportsAllDrives: true
  })

  const { body } = await fetch({
    hostname: 'www.googleapis.com',
    path: `/drive/v3/files/${fileId}?${params.toString()}`,
    method: 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json'
    }
  })
  if (!body?.modifiedTime) return null
  return new Date(body.modifiedTime)
}
