SQL2XLS
=======

Queries a SQLExpress server instance and exports the result into a spreadsheet.

### Usage

`sql2xls /path/to/config.json`

Where `config.json` is something like this:

```json
{
	"username": "user",
	"password": "pass",
	"server"  : "localhost",
	"database": "SIGA",
	"output"  : "Customers.xls",
	"query"   : "SELECT * FROM Customers"
}
```

This will execute `SELECT * FROM Customers` and output the result to `Customers.xls` (relative to `config.json`). If omitted it will be set to `config.xls` (same name as the `.json` but with different extension)

If `query` is omitted, it will try to read `config.sql` instead (same name as the `.json` but with different extension). You can also specify which file to read by setting the `read` parameter.

Any parameter can be overwritten by passing an argument. For instance, executing `query2xls /path/to/config.json --output another-file.xls` will execute the config file exactly but with a different output file.