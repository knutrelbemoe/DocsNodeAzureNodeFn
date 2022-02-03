const getGraphUser = require('./graph/get-graph-user');
const getGraphOrg = require('./graph/get-graph-org');
const getGraphRoot = require('./graph/get-graph-root');
const HTTPError = require('../lib/http-error');

module.exports = async (context, req) => {
  const token = req.headers.authorization || null
  if (!token) {
    throw new HTTPError(401, 'Unauthorized. Missing authorization header.')
  }

  let samAccountName, graphUser, graphOrg, orgRoot, myProfileProp

  try {
    graphUser = await getGraphUser(context, {
      token
    });

    myProfileProp = {
      user: graphUser
    };

  } catch (err) {
    context.log.error(['tasks', 'graph-user', 'err', err])
    throw new HTTPError(500, 'Unable to get user.')
  }

  try {
    orgRoot = await getGraphRoot(context, {
      token
    });

    myProfileProp = {
      ...myProfileProp,
      root: orgRoot
    };

  } catch (err) {
    context.log.error(['tasks', 'graph-user', 'err', err])
    throw new HTTPError(500, 'Unable to get user.')
  }

  try {
    graphOrg = await getGraphOrg(context, token)
    const userTenantId = graphOrg.value[0].id

    myProfileProp = {
      ...myProfileProp,
      organization: graphOrg
    };
    // if (userTenantId !== tenantId) {
    //   context.log.error(['tasks', samAccountName, `Tenant ID not matching ${tenantId}`, userTenantId])
    //   throw new HTTPError(401, 'Invalid tenant ID')
    // }
  } catch (err) {
    context.log.error(['tasks', graphUser, 'graph-org', 'err', err])
    throw new HTTPError(500, 'Unable to get organization')
  }

  return myProfileProp;
}