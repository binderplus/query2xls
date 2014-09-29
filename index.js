#!/usr/bin/env node
var fs          = require('fs')
var path        = require('path')
var tedious     = require('tedious')
var yargs       = require('yargs')
var async       = require('async')
var excel       = require('excel-export')

// Parse CLI
var argv = yargs
    .usage('query2xls --c path/to/config.json')

    // Query Related INPUT/OUTPUT
    .alias('q', 'query')    .describe('q', 'SQL Query String')        
    .alias('r', 'read')     .describe('r', 'Read SQL Query File (UTF8)')   
    .alias('o', 'output')   .describe('o', 'Output XLS File')

    // Database Settings
    .alias('u', 'username') .describe('u', 'Database Username')
    .alias('p', 'password') .describe('p', 'Database Password')
    .alias('s', 'server')   .describe('s', 'Database Server Address')
    .alias('i', 'instance') .describe('i', 'Database Instance Name')
    .alias('d', 'database') .describe('d', 'Database Name')

    // Config File
    .alias('c', 'config')   .describe('c', 'Configuration file.')
    .argv

// Read config file
var config = {}
if (!!argv.config) {
    if (!fs.existsSync(argv.config)) {
        console.error('ERROR: Can\'t find config file ' + argv.config + '\r\n')
        yargs.showHelp()
        process.exit(1)
    }
    config = JSON.parse(fs.readFileSync(argv.config, "utf8"))

    // Set output file path relative to config file
    config.output = path.resolve(path.dirname(argv.config), config.output)
}

// Apply defaults
config.server   = config.server   || 'localhost'
config.output   = config.output   || 'output.xls'
config.instance = config.instance || 'SQLExpress'

// Overwrite with arguments
if (!!argv.query)    config.query    = argv.query
if (!!argv.read)     config.read     = argv.read
if (!!argv.output)   config.output   = argv.output
if (!!argv.username) config.username = argv.username
if (!!argv.password) config.password = argv.password
if (!!argv.server)   config.server   = argv.server
if (!!argv.instance) config.instance = argv.instance
if (!!argv.database) config.instance = argv.database

// Demand query, or query file
if (!config.query) {
    if (!config.read) {
        console.error('ERROR: Either query string or query file is required!\r\n')
        yargs.showHelp()
        process.exit(1)
    }
    if (!fs.existsSync(config.read)) {
        console.error('ERROR: Can\'t find queryFile ' + config.read + '\r\n')
        yargs.showHelp()
        process.exit(1)
    }

    // Read from file
    config.query = fs.readFileSync(config.read, "utf8")
}

// Demand arguments
var demand = [ 'query', 'output', 'username', 'password', 'instance', 'database' ]
demand.forEach(function (arg) {
    if (!config.hasOwnProperty(arg) || config[arg] === 'undefined') {
        console.error('ERROR: ' + arg + ' is required!\r\n')
        yargs.showHelp()
        process.exit(1)
    }
})

// Connect to database
var connection = new tedious.Connection({
    userName: config.username,
    password: config.password,
    server  : config.server,
    options : {
        instanceName: config.instance,
        databaseName: config.database
    }
}).on('connect', function (err) {
    if (!!err) throw err
    connection.connected = true
    console.log('Established database connection!')
})

// Work
connection.on('connect', function () {

    var columnNames = null
    var data = []

    var request = new tedious.Request(config.query, function (err) {
        if (!!err) throw err;

        // Finished, do something with data
        var conf = {}

        // Column metadata
        conf.cols = columnNames.map(function (col) {
            return {
                caption: col,
                type: 'string'
            }
        })

        // Rows
        conf.rows = data

        console.log('\r\nSaving file!\r\n')
        fs.writeFileSync(config.output, excel.execute(conf), 'binary')
        process.exit()
    })

    request.on('row', function (columns) {
        // Populate Column Metadata
        if (!columnNames) {
            columnNames = columns.map(function (col) {
                return col.metadata.colName
            })
        }

        // Row values
        var row = columns.map(function (col) {
            return col.value
        })

        data.push(row)

        // Show progress
        console.log('.')
    })

    connection.execSql(request)
})
