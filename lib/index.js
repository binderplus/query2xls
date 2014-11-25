var tedious     = require('tedious')
var xlsx        = require('node-xlsx')
var _ = {   // Lodash
    defaults    : require('lodash.defaults')
}

module.exports  = function SQL2Xls (config, callback, logFn) {

    // Config & Defaults
    var config = _.defaults(config || {}, {
        server          : 'localhost', 
        instance        : 'SQLExpress',
        hideColumnNames : false
    })

    // Config Sanity Checks
    var demand = ['query', 'username', 'password', 'instance', 'database']
    .forEach(function (arg) {
        if (!config.hasOwnProperty(arg) || typeof config[arg] !== 'string') {
            throw new TypeError('config.'+arg+' is required and expected to be a string.')
        }
    })

    // Callback is also required
    if (typeof callback !== 'function')
        throw new TypeError('callback is required and expected to be a function')

    // Default logFn
    if (typeof logFn === 'undefined') logFn = function() {}
    else if (typeof logFn !== 'function') throw new TypeError('logFn expected to be a function')

    // Connect to database
    var connection = new tedious.Connection({
        userName: config.username,
        password: config.password,
        server  : config.server,
        options : {
            instanceName: config.instance,
            databaseName: config.database
        }

    // Work
    }).on('connect', function (err) {
        if (!!err) throw err
        logFn('Established database connection!')

        var columnNames = null
        var data = []

        var request = new tedious.Request(config.query, function (err) {
            if (!!err) throw err;

            // Prepend Names
            if (!config.hideColNames) data.unshift(columnNames)

            logFn('Saving file!')
            callback(null, xlsx.build([{ data: data }]))
        })

        request.on('row', function (columns) {

            // Populate Column Metadata
            if (!columnNames)
                columnNames = columns.map(function (col) { return col.metadata.colName })

            // Row values
            var row = columns.map(function (col) { return col.value })

            data.push(row)

            // Log progress
            if (data.length%100 == 0) 
                console.log('Receiving row #', Math.floor(data.length/100)*100)
        })

        connection.execSql(request)
    })
}