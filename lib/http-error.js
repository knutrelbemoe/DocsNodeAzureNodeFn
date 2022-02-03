const { STATUS_CODES } = require('http')
const uppercamelcase = require('uppercamelcase')

// Stolen from @TooTallNate - https://gist.github.com/TooTallNate/4fd641f820e1325695487dfd883e5285

class HTTPError extends Error {
  constructor (code, message) {
    super(message || STATUS_CODES[code])

    this.name = toName(code)
    this.statusCode = code
  }
}

/**
 * Converts an HTTP status code to an Error `name`.
 * Ex:
 *   302 => "Found"
 *   404 => "NotFoundError"
 *   500 => "InternalServerError"
 */
const toName = (code) => {
  const suffix = (code / 100 | 0) === 4 || (code / 100 | 0) === 5 ? 'error' : ''
  return uppercamelcase(String(STATUS_CODES[code]).replace(/error$/i, ''), suffix)
}

module.exports = HTTPError
