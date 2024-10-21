# sql_lib
### Вариант использования 1
```javascript
var path = "x-local://wt/web" + "путь до библиотеки" + "sql_lib.js"
//DropFormsCache(path)
var lib = OpenCodeLib(path)

var connection = "строка ODBC подключения"
var ssql = "SELECT * FROM таблица"
var result = lib.optXExec(ssql, connection)
// в reult упадет результат выборки ssql
```
