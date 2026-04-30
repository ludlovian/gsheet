import https from 'node:https'

export async function fetch ({ json = true, ...opts }) {
  opts.headers ??= {}

  if (opts.body && typeof opts.body === 'object') {
    opts.body = JSON.stringify(opts.body)
    /* c8 ignore start */
    if (opts.method === 'POST' && !opts.headers['Content-Type']) {
      opts.headers['Content-Type'] = 'application/json'
    }
    /* c8 ignore stop */
    opts.headers['Content-Length'] = Buffer.byteLength(opts.body)
  }

  const response = await doRequest(opts)
  const { statusCode } = response
  response.ok = statusCode >= 200 && statusCode <= 299
  /* c8 ignore start */
  if (!response.ok) {
    const { statusMessage } = response
    const err = new Error(`Fetch failed [${statusCode}]: ${statusMessage}`)
    const { body, headers, ...rest } = opts
    err.request = rest
    throw err
  }
  /* c8 ignore stop */

  response.body = ''
  response.setEncoding = 'utf8'
  for await (const chunk of response) {
    response.body += chunk
  }

  if (json && response.body) {
    response.body = JSON.parse(response.body)
  }
  return response
}

function doRequest ({ body, ...rest }) {
  return new Promise((resolve, reject) => {
    const req = https.request(rest, resolve)
    req.on('error', reject)
    if (body) req.write(body)
    req.end()
  })
}
