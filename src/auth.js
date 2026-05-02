import assert from 'node:assert'
import { readFileSync, existsSync } from 'node:fs'
import crypto from 'node:crypto'
import * as jeeves from '@ludlovian/jeeves'

const CREDENTIALS_FILE = 'creds/credentials.json'
assert(existsSync(CREDENTIALS_FILE), `${CREDENTIALS_FILE} is missing`)
const credentials = JSON.parse(readFileSync(CREDENTIALS_FILE, 'utf8'))

const tokenMap = new Map() // scope => { token, expiryMs }

export async function getAccessToken (scope) {
  /* c8 ignore next */
  if (Array.isArray(scope)) scope = scope.join(' ')
  const entry = tokenMap.get(scope)
  if (entry) {
    if (entry.expiryMs > Date.now()) return entry.token
  }

  const jwt = getSignedJWT(scope)

  const url = 'https://oauth2.googleapis.com/token'
  const headers = { accept: 'application/json' }
  const body = new URLSearchParams({
    grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
    assertion: jwt
  })

  const resp = await jeeves.post(url, { headers, body })

  const { access_token: token } = await resp.json()
  const expiryMs = Date.now() + 55 * 60 * 1e3 // 55 mins
  tokenMap.set(scope, { token, expiryMs })
  return token
}

function getSignedJWT (scope) {
  const nowSec = Math.floor(Date.now() / 1000)

  const header = { alg: 'RS256', typ: 'JWT' }

  const claimSet = {
    iss: credentials.client_email,
    scope,
    aud: 'https://oauth2.googleapis.com/token',
    exp: nowSec + 3600,
    iat: nowSec
  }

  // 2. Create the Signature Input
  const signatureInput = `${b64(header)}.${b64(claimSet)}`

  // 3. Sign with the Private Key
  const signer = crypto.createSign('RSA-SHA256')
  signer.update(signatureInput)
  const signature = signer
    .sign(credentials.private_key, 'base64')
    .replace(/=/g, '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')

  return `${signatureInput}.${signature}`
}

function b64 (obj) {
  return Buffer.from(JSON.stringify(obj))
    .toString('base64')
    .replace(/=/g, '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
}
