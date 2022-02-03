const getGraphToken = require('./get-graph-token')
const axios = require('axios')
const { graph: { org } } = require('../../config')

module.exports = async (context, token) => {
  if (!token) {
    // If no token is provided - gather token
    context.log(['tasks', 'get-graph-org', 'get-graph-token'])
    token = await getGraphToken(context)
    context.log(['tasks', 'get-graph-org', 'get-graph-token', 'length', token.length])
  }

  const options = {
    url: org.url,
    method: 'GET',
    headers: {
      Authorization: token
    },
    params: {
      $select: org.properties
    }
  }

  try {
    const { data } = await axios(options)
    return data
  } catch (err) {
    context.log.error(['tasks', 'get-graph-org', 'err', err])
    throw err
  }
}
