const axios = require('axios')
const qs = require('qs')
const { graph: { auth } } = require('../../config')

const NodeCache = require('node-cache')
const cache = new NodeCache({ stdTTL: 3000 })

module.exports = async (context) => {
  if (cache.get('graphToken')) {
    context.log(['tasks', 'get-graph-token', 'return cached token'])
    return cache.get('graphToken')
  }

  const authOptions = {
    client_id: auth.clientId,
    scope: auth.scope,
    client_secret: auth.secret,
    grant_type: auth.grantType
  }

  const options = {
    url: auth.url,
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    data: qs.stringify(authOptions)
  }

  try {
    const { data } = await axios(options)
    const token = `${data.token_type} ${data.access_token}`.trim()
    cache.set('graphToken', token, data.expires_in)

    return token
  } catch (err) {
    context.log.error(['tasks', 'get-graph-token', 'err', err])
    throw err
  }
}
