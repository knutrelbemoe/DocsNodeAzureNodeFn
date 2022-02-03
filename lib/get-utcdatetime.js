const moment = require('moment');

module.exports = (dateadd) => {

    if (isNaN(dateadd)) {
        let currentDatetimeUtc = moment();
        return currentDatetimeUtc.format(moment.HTML5_FMT.DATE);
    } else {
        let DatetimeUtc = moment();
        let nextDatetime = DatetimeUtc.add(dateadd, 'days').format(moment.HTML5_FMT.DATE);
        return nextDatetime;
    }

}