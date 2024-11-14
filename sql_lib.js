/**
 * Получить строку подключения из конфигурационного файла с estaff сервера
 * @returns {string}
 */
function getStrFromEstaffServer() {
    try {
        // работает только с новым e-staff, где пароли кэшируются
        var path = FilePath(AppDirectoryPath(), "app_config.xml");

        if (!FilePathExists(path)) {
            return "";
        }
        var card = OpenDoc(FilePathToUrl(path), "form=//app/sx_app_config.xmd").TopElem;

        // собираем строку подключения
        return "Provider=SQLOLEDB.1;" +
                "User ID=" + card.storage.login + ";" +
                "Password=" + StrStdDecrypt(card.storage.password_ed) + ";" +
                "Trusted_Connection=False;" +
                "Database=" + card.storage.database + ";" +
                "Server=" + card.storage.server + ";";
    } catch (err) {}

    return "";
}


/**
 * Соединить все поля объекта через разделитель
 * @param {object} obj - объект с параметрами подключения
 * @param {string} separator - разделитель
 * @returns {string}
 */
function joinParams(obj, separator) {
    var result = [];

    var field;
    for (field in obj) {
        result.push(field + "=" + obj.GetOptProperty(field));
    }

    return result.join(separator);
}


/**
 * Получить строку подключения из кастомного конфигурационного файла
 * @param {string} serverName - имя сервера
 * @returns {string}
 */
function getStrFromConfig(serverName) {
    try {
        var pathToConfig = FilePath(AppDirectoryPath(), "server.config");
        if (!FilePathExists(pathToConfig)) {
            throw new Error("server.config not found");
        }

        var paramsConnection = ParseJson(LoadFileData(pathToConfig)).GetOptProperty("connections");
        var paramsServer = paramsConnection.GetOptProperty(serverName);

        if (paramsServer != undefined) {
            return joinParams(paramsServer, ";");
        }
    } catch (err) {
        alert("ERROR sql_lib.js: " + err);
    }

    return undefined;
}


/**
 * Получить строку подключения к БД
 * @param {any} sConnection - строка подключения
 * @returns {string}
 */
function getConnectionString(sConnection) {
    var result = getStrFromConfig(sConnection);
    if (result != undefined) {
        return result;
    }

    if (sConnection != undefined && sConnection != "") {
        return sConnection;
    }

    return getStrFromEstaffServer();
}


/**
 * Установить соединение с БД
 * @param {any} connection - подключение
 * @returns {object}
 */
function getActiveConnection(connection) {
    // если подключение уже существует
    if (DataType(connection) == "object" && connection.GetOptProperty("ADOConnect") != undefined) {
        return {
            success: true,
            activeConnection: connection.ADOConnect,
        };
    }

    // получить строку подключения
    var ConnectionString = getConnectionString(connection);
    if (ConnectionString == "") {
        return {
            success: false,
            error: "app_config.xml not found",
        };
    }

    // создать подключение
    var ADOConnection = new ActiveXObject("ADODB.Connection");
    try {
        ADOConnection.Open(ConnectionString);
    } catch (err) {
        return {
            success: false,
            error: err + "",
        };
    }

    return {
        success: true,
        activeConnection: ADOConnection,
    };
}


/**
 * Установить параметры команды
 * @param {object} param - параметры
 * @returns {object}
 */
function setCommand(param) {
    var ADOCommand = new ActiveXObject("ADODB.Command");

    ADOCommand.ActiveConnection = param.activeConnection;
    ADOCommand.CommandType = 1;
    ADOCommand.CommandTimeout = 3600;
    ADOCommand.CommandText = param.ssql;

    return ADOCommand;
}


/**
 * Получение набора записей из БД
 * @param {object} command - команда
 * @returns {object} - В один момент времени объект может ссылаться только на одну запись
 */
function getRecords(command) {
    var ADORecordSet = new ActiveXObject("ADODB.RecordSet");

    try {
        ADORecordSet.Open(command);
    } catch (err) {
        return {
            success: false,
            error: err,
        };
    }

    return {
        success: true,
        ADORecord: ADORecordSet,
    };
}


/**
 * Преобразовываем объект выборки в массив c объектами
 * @param {object} record - объект выборки
 * @returns {Array}
 */
function convertRecordsToArray(record) {
    var result = [];

    var obj, i;
    var countFields = record.Fields.Count;

    while (!record.EOF) {
        if (record.BOF) {
            record.MoveNext();
            continue;
        }

        obj = {};
        for (i = 0; i < countFields; i++) {
            obj[record.Fields.Item(i).Name] = record.Fields.Item(i).Value;
        }

        result.push(obj);
        record.MoveNext();
    }

    return result;
}


/**
 * Преобразовать объукт выборки в объукт с полями field
 * @param {object} record - объект выборки
 * @param {object} field - поле, согласно которому будет группироваться запись
 * @returns {object}
 */
function convertRecordsToObject(record, field) {
    var result = {};
    var sField = String(field);
    var sValueForField;

    var countFields = record.Fields.Count;
    while (!record.EOF) {
        if (record.BOF) {
            record.MoveNext();
            continue;
        }

        sValueForField = String(record.Fields(sField).Value);
        result[sValueForField] = {};
        for (i = 0; i < countFields; i++) {
            result[sValueForField][record.Fields.Item(i).Name] = record.Fields.Item(i).Value;
        }

        record.MoveNext();
    }

    return result;
}


/**
 * Проверить, нужен ли нам результирующий объект
 * @param {object} param - параметры для результирующего объекта
 * @returns {boolean}
 */
function isObjectResult(param) {
    if (param == undefined) {
        return false;
    }

    if (ObjectType(param) != "JsObject") {
        return false;
    }

    if (param.GetOptProperty("type") != "object") {
        return false;
    }

    if (param.GetOptProperty("field", "") == "") {
        return false;
    }

    return true;
}


/**
 * Преобразование объекта выборки в данные согласно param
 * @param {object} record - объект выборки
 * @param {object} param - параметры
 * @returns {any}
 */
function convertRecords(record, param) {
    if (isObjectResult(param)) {
        try {
            // преобразовать объект выборки в объект с полями field
            return convertRecordsToObject(record, param.field);
        } catch (err) {
            // скорее всего сюда упадет ошибка, если мы передали некорректное поле из таблицы
            // обрабатываем что бы  функционал вернул массив, а не упал в ошибку
            alert("ERROR: " + err);
        }
    }

    // преобразовывает объект выборки в стандартный массив объектов
    return convertRecordsToArray(record);
}


/**
 * Выполнить sql запрос
 * @param {string} sql - sql запрос
 * @param {any} connection - строка подключения
 * @returns {object}
 */
function exec(sql, connection) {
    var connect = getActiveConnection(connection);
    if (!connect.success) {
        return {
            successfull: false,
            error: connect.error + "",
        };
    }

    var command = setCommand({
        activeConnection: connect.activeConnection,
        ssql: sql,
    });

    var records = getRecords(command);
    if (!records.success) {
        return {
            successfull: false,
            error: records.error + "",
            ADOConnect: connect.activeConnection,
        };
    }

    return {
        successfull: true,
        ADORecord: records.ADORecord,
        ADOConnect: connect.activeConnection,
    };
}


/**
 * Выполнить sql запрос и вернуть массив объектов
 * @param {string} sql - sql-запрос
 * @param {any} connection - подключение к бд
 * @param {object} param - в этих параметрах мы задаем какой тип нам вернуть
 * @returns {array || object}
 */
function optXExec(sql, connection, param) {
    var resExec = exec(sql, connection); // получить данные
    if (!resExec.successfull) {
        // закрыть соединение если оно открыто
        if (resExec.GetOptProperty("ADOConnect") != undefined) {
            resExec.ADOConnect.close();
        }

        return [];
    }

    var result = convertRecords(resExec.ADORecord, param);
    resExec.ADOConnect.close(); // закрыть соединение

    return result;
}


/**
 * Для linux можно посмотреть в этом направлении
 *
 * смотри wtv_export_odbc.js
 * 
 * var oAssembly = tools.dotnet_host.Object.GetAssembly('Websoft.Discharge.dll');
 * db = oAssembly.CreateClassObject('Websoft.Discharge.Discharge');*
 *
 *
 *
 *

Коллеги, привествую.
Вопрос по подключению к внешней базе oracle для систем на линукс.
Посмотрел файлик wtv_export_odbc.js
  var oAssembly = tools.dotnet_host.Object.GetAssembly('Websoft.Discharge.dll');
  db = oAssembly.CreateClassObject('Websoft.Discharge.Discharge');
  db.Init( _source.source_name.Value, _source.db_type.Value );
  db.SqlExecute( _table.select_query.Value );
  db.GetNext();
  var Json = db.GetRowJson();
В целом понятно , но вопрос, а если вернуть надо не строковое значение а бинарный объект?
Допустим есть столбец "FileData" типа RAW содержащий данные файла.
а db.GetRowJson в этом случае возвращает строку  System.Byte[]
каким методом можно вернуть в нормальном виде.
и есть ли вообще список доступных методов

 */
