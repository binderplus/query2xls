SQL Query 2 XLS
===============

Queries a SQLExpress server instance and exports the result into a spreadsheet.

### Usage

`query2xls --c /path/to/config.json`

Where `config.json` is something like this:

```json
{
	"username": "user",
	"password": "pass",
	"server"  : "localhost",
	"database": "SIGA",
	"output"  : "./output/Customers.xls",
	"query"   : "SELECT * FROM Customers"
}
```

This will execute `SELECT * FROM Customers` and output the result to `./output/Customers.xls` (relative to `config.json`).

All parameters can be overwritten by passing an argument. For instance, executing `query2xls --c /path/to/config.json --output another-file.xls` will execute the config file exactly but with a different output file.