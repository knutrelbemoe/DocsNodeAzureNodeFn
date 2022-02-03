const axios = require('axios');
const {
  graph: {
    calender
  }
} = require('../../config');

module.exports = async (context, token, userPrincipalName) => {
  const options = {
    url: `${calender.graphUrl}`,
    method: 'GET',
    headers: {
      Authorization: token
    },
    params: {
      $limit: calender.maxTasks,
      $filter: calender.filter,
      $orderby: calender.orderBy
    }
  }

  try {
    const {
      data
    } = await axios(options);

    return data;
  } catch (err) {
    context.log.error(['tasks', 'get-outlook-tasks', 'err', err]);
    throw err;
  }
}