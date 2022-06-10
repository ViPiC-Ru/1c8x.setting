/* 0.1.0 настройка конфигурационного файла клиента 1с по данным из active directory

cscript 1c8x.setting.min.js <location> [<prefix>] [<map>...] \\ [<config>...]

<location>      - Путь к конфигурационному файлу для создания, изменения или удаления.
<prefix>        - Префикс групп пользователя для получения параметров из их атрибутов.
<map>           - Соответствие параметров конфигурации и атрибутов групп пользователя.
                  Если не указаны, то настройки применятся только если пользователь
                  состоит в группах, удовлетворяющих префиксу.
<config>        - Параметры конфигурации и их значения. Пустое значение для удаления.

*/

var setting = new App({
    argWrap: '"',                           // основное обрамление аргументов
    keyDelim: "=",                          // разделитель ключа от значения
    lineDelim: "\r\n",                      // разделитель строк
    putDelim: "\\\\",                       // разделитель потоков параметров
    adProtocol: "LDAP:",                    // протокол для подключения к active directory
    intСharset: "unicode",                  // внутренняя кодировка
    subFilePath: "..\\1cv8\\1cv8strt.pfl",  // дополнительный файл со специальными настройками 
    subFileСharset: "utf-8"                 // кодировка дополнительного файла
});

// подключаем зависимые свойства приложения
(function (wsh, app, undefined) {
    app.lib.extend(app, {
        fun: {// зависимые функции частного назначения
        },
        init: function () {// функция инициализации приложения
            var key, value, index, length, list, mainFileLocation, subFileLocation, prefix,
                shell, fso, adsi, user, userDN, group, groupDN, input, output, isAllowChange,
                left, center, right, lines, path, attribute, before, after, isDelim,
                map = {}, config = {}, special = {}, control = {}, error = 0;

            adsi = new ActiveXObject("ADSystemInfo");
            shell = new ActiveXObject("WScript.Shell");
            fso = new ActiveXObject("Scripting.FileSystemObject");
            isAllowChange = true;// разрешено ли внести изменения в файлы
            // получаем стандартные параметры
            if (!error) {// если нет ошибок
                length = wsh.arguments.length;// получаем длину
                for (index = 0; index < length; index++) {// пробигаемся по параметрам
                    value = wsh.arguments.item(index);// получаем очередное значение
                    // путь к файлу конфигурации
                    key = "location";// ключ проверяемого параметра
                    if (0 == index) {// если первый параметр
                        path = shell.expandEnvironmentStrings(value);
                        path = fso.getAbsolutePathName(path);
                        // присваиваем значение
                        mainFileLocation = path;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                    // префик групп для загрузки параметров из ad
                    key = "prefix";// ключ проверяемого параметра
                    if (1 == index && app.val.putDelim != value) {
                        // присваиваем значение
                        prefix = value;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                    // если закончились стандартные параметры
                    break;// остававливаем получние параметров
                };
            };
            // проверяем обязательные параметры
            if (!error) {// если нет ошибок
                if (mainFileLocation) {// если проверка пройдена
                } else error = 1;
            };
            // получаем параметры маппинга
            if (!error) {// если нет ошибок
                isDelim = false;// сбрасываем значение
                while (index < length && !isDelim && !error) {
                    value = wsh.arguments.item(index);// получаем значение
                    if (app.val.putDelim != value) {// если не разделитель потоков
                        key = app.lib.strim(value, null, app.val.keyDelim, false, false);
                        if (key) {// если параметр имеет нужный формат
                            value = app.lib.strim(value, app.val.keyDelim, null, false, false);
                            list = value.split(app.val.argWrap);// вспомогательная переменная
                            if (3 == list.length && !list[0] && !list[2]) value = list[1];
                            if (value) {// если задано не пустое значение
                                map[key] = value;
                            } else error = 3;
                        } else error = 2;
                    } else isDelim = true;
                    index++;
                };
            };
            // получаем параметры конфигурации
            if (!error) {// если нет ошибок
                while (index < length && !error) {
                    value = wsh.arguments.item(index);// получаем значение
                    key = app.lib.strim(value, null, app.val.keyDelim, false, false) || value;
                    if (key) {// если параметр имеет нужный формат
                        value = app.lib.strim(value, app.val.keyDelim, null, false, false);
                        list = value.split(app.val.argWrap);// вспомогательная переменная
                        if (3 == list.length && !list[0] && !list[2]) value = list[1];
                        value = shell.expandEnvironmentStrings(value);
                        if (!(key in config)) config[key] = [];
                        if (value) config[key].push(value);
                    } else error = 4;
                    index++;
                };
            };
            // работаем в зависимости от параметров для active directory
            if (prefix) {// если задан префикс
                // получаем информацию о пользователе
                if (!error) {// если нет ошибок
                    try {// пробуем выполнить комманду
                        userDN = adsi.userName;
                        user = GetObject(app.val.adProtocol + "//" + userDN);
                    } catch (e) {// если ошибка
                        error = 5;
                    };
                };
                // получаем параметры конфигурации из групп пользователя
                if (!error) {// если нет ошибок
                    if (!app.lib.count(map)) isAllowChange = false;
                    try {// пробуем выполнить комманду
                        list = user.get("memberOf").toArray();
                        for (var i = 0, iLen = list.length; i < iLen; i++) {
                            groupDN = list[i];// получаем очередной идентификатор
                            if (3 == groupDN.toLowerCase().indexOf(prefix.toLowerCase())) {
                                isAllowChange = true;// разрешено ли внести изменения в файлы
                                if (app.lib.count(map)) {// если есть параметры для меппинга
                                    group = GetObject(app.val.adProtocol + "//" + groupDN);
                                    for (var key in map) {// пробигаемся по параметрам меппинга
                                        attribute = map[key];
                                        value = group.get(attribute);
                                        value = shell.expandEnvironmentStrings(value);
                                        if (!(key in config)) config[key] = [];
                                        if (value) config[key].push(value);
                                    };
                                };
                            };
                        };
                    } catch (e) {// если ошибка
                        error = 6;
                    };
                };
            };
            // переносим специальные параметры из конфигурации и отмечаем проверки
            if (!error) {// если нет ошибок
                for (var key in config) {// пробигаемся по параметрам
                    switch (key) {// специализированные параметры
                        case "LRInfoBaseIDListSize":    // сколько запоминать последних
                        case "ShowIBsAsTree":           // отображать в виде дерева
                        case "AutoSortIBs":             // сортировать по наименованию
                        case "ShowRecentIBs":           // отображать недавние
                        case "DefaultConnectionSpeed":  // скорость подключения
                            if (1 == config[key].length) {
                                value = config[key][0];
                                special[key] = value;
                                delete config[key];
                            } else error = 7;
                            break;
                        default:// параметр конфигурации
                            control[key] = true;
                    };
                };
            };
            // работаем в зависимости от разрешения на внесение изменений
            if (isAllowChange) {// если разрешено вносить изменения
                // проверяем существование дополнительного файла 
                if (!error && app.lib.count(special)) {// если нужно выполнить
                    path = fso.getParentFolderName(mainFileLocation);
                    path = fso.buildPath(path, app.val.subFilePath);
                    path = fso.getAbsolutePathName(path);
                    if (fso.fileExists(path)) subFileLocation = path;
                };
                // работаем в зависимости от существования дополнительньного файла
                if (subFileLocation) {// если дополнительньный файл существует
                    // читаем данные дополнительньного файла в строки
                    if (!error) {// если нет ошибок
                        input = app.wsh.getFileText(subFileLocation);
                        input = app.wsh.iconv(app.val.subFileСharset, app.val.intСharset, input);
                        input = input.substring(3);// отсекаем BOM
                        lines = input.split(app.val.lineDelim);
                    };
                    // изменяем данные в строках
                    if (!error) {// если нет ошибок
                        left = '",'; center = '},"'; right = '",';// разделители
                        for (var i = 0, iLen = lines.length - 1; i < iLen; i++) {
                            for (var key in special) {// пробигаемся по параметрам
                                if (~lines[i].indexOf(center + key + right)) {// если найден параметр
                                    value = app.lib.strim(lines[i + 1], left, center, false, false);
                                    if (value && value != special[key]) {// если нужно изменить значение
                                        before = left + value + center;// что заменить
                                        after = left + special[key] + center;// на что заменить
                                        lines[i + 1] = lines[i + 1].replace(before, after);
                                    };
                                };
                            };
                        };
                    };
                    // записываем данные из строк в дополнительньный файл
                    if (!error) {// если нет ошибок
                        output = lines.join(app.val.lineDelim);
                        if (input != output) {// если требуется внести изменения
                            output = app.wsh.iconv(app.val.intСharset, app.val.subFileСharset, output);
                            if (app.wsh.setFileText(subFileLocation, output, false)) {
                            } else error = 8;
                        };
                    };
                };
                // читаем данные основного файла в строки
                if (!error) {// если нет ошибок
                    input = app.wsh.getFileText(mainFileLocation);
                    lines = input.split(app.val.lineDelim);
                };
                // переносим данные из строк в конфигурацию
                if (!error) {// если нет ошибок
                    for (var i = 0, iLen = lines.length; i < iLen; i++) {
                        value = lines[i];// получаем очередную строку
                        key = app.lib.strim(value, null, app.val.keyDelim, false, false);
                        if (key) {// если параметр имеет нужный формат
                            value = app.lib.strim(value, app.val.keyDelim, null, false, false);
                            if (!control[key]) {// если параметр не контролируется
                                if (!(key in config)) config[key] = [];
                                if (value) config[key].push(value);
                            };
                        };
                    };
                };
                // переносим данные из конфигурации в строки
                if (!error) {// если нет ошибок
                    lines = [];// сбрасываем значение
                    for (var key in config) {// пробигаемся по параметрам
                        for (var i = 0, iLen = config[key].length; i < iLen; i++) {
                            value = config[key][i];// получаем очередное значение
                            value = key + app.val.keyDelim + value;
                            lines.push(value);
                        };
                    };
                };
                // записываем данные из строк в основной файл
                if (!error) {// если нет ошибок
                    output = lines.join(app.val.lineDelim);
                    if (input != output) {// если требуется внести изменения
                        if (output) {// если нужно создать или изменить файл
                            path = fso.getParentFolderName(mainFileLocation);
                            if (app.wsh.getFolder(path, true)) {
                                if (app.wsh.setFileText(mainFileLocation, output, false)) {
                                } else error = 10;
                            } else error = 9;
                        }else{// если нужно удалить файл
                            path = mainFileLocation;// путь к файлу
                            if (fso.fileExists(path)){// если файл существует
                                try {// пробуем выполнить комманду
                                    fso.deleteFile(path ,true);
                                } catch (e) {// если ошибка
                                    error = 11;
                                };
                            };
                        };
                    };
                };
            };
            // завершаем сценарий кодом
            wsh.quit(error);
        }
    });
})(WSH, setting);
// запускаем инициализацию
setting.init();