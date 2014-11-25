#!/usr/bin/env node
var fs          = require('fs')
var path        = require('path')
var yargs       = require('yargs')
var sql2xls		= require('./index.js')

// Parse CLI
var argv = yargs
    .usage('query2xls path/to/config.json')

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
    .argv

// Read config file
var configFile = argv._[0]
if (typeof configFile !== 'string' && !fs.existsSync(configFile)) {
    console.error('ERROR: Can\'t find config file ' + configFile + '\r\n')
    yargs.showHelp()
    process.exit(1)
}

// Read config
config = JSON.parse(fs.readFileSync(configFile, "utf8"))

// Set output file path relative to config file
config.output = path.resolve(path.dirname(configFile), 
	config.output || path.basename(configFile, path.extname(configFile))+'.xls')

// Set queryFile relative to config file
config.read = path.resolve(path.dirname(configFile), 
    config.read || path.basename(configFile, path.extname(configFile))+'.sql')

// Overwrite with arguments
if (!!argv.query)    config.query    = argv.query
if (!!argv.read)     config.read     = argv.read
if (!!argv.output)   config.output   = argv.output
if (!!argv.username) config.username = argv.username
if (!!argv.password) config.password = argv.password
if (!!argv.server)   config.server   = argv.server
if (!!argv.instance) config.instance = argv.instance
if (!!argv.database) config.database = argv.database

// Apply defaults
config.output = config.output || 'output.xls'

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

// Execute
sql2xls(config, function (err, buffer) { 
	if (!!err) throw err;
	fs.writeFileSync(config.output, buffer)
	process.exit(0)
}, console.log)