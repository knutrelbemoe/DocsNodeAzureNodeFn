const getGraphToken = require('./get-graph-token')
const axios = require('axios')
const { graph: { user } } = require('../../config')

module.exports = async (context, { token, userPrincipalName }) => {
  if (!token) {
    // If no token is provided - gather application graph token
    context.log(['tasks', 'get-graph-user', userPrincipalName, 'get-graph-token'])
    token = await getGraphToken(context)
    context.log(['tasks', 'get-graph-user', userPrincipalName, 'get-graph-token', 'length', token.length])
  }

  const options = {
    url: userPrincipalName ? `${user.userUrl}/${userPrincipalName}` : user.meUrl,
    method: 'GET',
    headers: {
      Authorization: token
    },
    params: {
      $select: user.properties
    }
  }

  try {
    context.log(['tasks', 'get-graph-user', userPrincipalName, 'get-user', 'url', options.url])

    const { data } = await axios(options)
    return data
  } catch (err) {
    context.log.error(['tasks', 'get-graph-user', 'err', err])
    throw err
  }
}
