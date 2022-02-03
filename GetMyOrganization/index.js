const authGraphUser = require('../lib/auth-graph-user')
const getUserTasks = require('../lib/get-user-calender')

module.exports = async (context, req) => {
    try {
        const graphUser = await authGraphUser(context, req)
        context.log(['tasks', 'get-my-profile', graphUser.user])


        context.res = {
            status: 200,
            body: {
                myOrgProfile: graphUser,

            }
        }
    } catch (err) {
        context.res = {
            status: err.statusCode || 500,
            body: err.message || 'Internal server error'
        }
    }
}