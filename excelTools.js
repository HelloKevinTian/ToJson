/**
 *  Excel to JSON
 *
 */

var xlsx = require('node-xlsx');
var fs = require('fs');
var path = require('path');
var log4js = require('log4js').getLogger('');
var async = require("async");
var util = require("util");
var ruleConfigMaker = require('./ruleConfigMaker');

/**
 * 检查创建文件夹
 *
 * @param  {String} url 指定文件夹路径
 * @param  {Number} mode 目录权限
 * @return {Void}
 */
function mkdirSync(url, mode) {

    var arr = url.split('/');

    mode = mode || 0755;

    if (arr[0] === '.') { //处理 ./aaa
        arr.shift();
    }
    if (arr[0] == '..') { //处理 ../ddd/d
        arr.splice(0, 2, arr[0] + '/' + arr[1])
    }

    function inner(cur) {
        if (!fs.existsSync(cur)) { //不存在就创建一个
            fs.mkdirSync(cur, mode)
        }
        if (arr.length) {
            inner(cur + '/' + arr.shift());
        }
    }
    arr.length && inner(arr.shift());
};

/**
 * json转换为字符串
 *
 * @param  {Object} 指定json
 * @return {String}
 */
function jsonToStr(a) {
    function j(a, b) {
        var c, d, g, k, l = e,
            m, n = b[a];
        n && typeof n == "object" && typeof n.toJSON == "function" && (n = n.toJSON(a));
        typeof h == "function" && (n = h.call(b, a, n));
        switch (typeof n) {
            case "string":
                return i(n);
            case "number":
                return isFinite(n) ? String(n) : "null";
            case "boolean":
            case "null":
                return String(n);
            case "object":
                if (!n) return "null";
                e += f;
                m = [];
                if (Object.prototype.toString.apply(n) === "[object Array]") {
                    k = n.length;
                    for (c = 0; c < k; c += 1) m[c] = j(c, n) || "null";
                    g = m.length === 0 ? "[]" : e ? "[\n" + e + m.join(",\n" + e) + "\n" + l + "]" : "[" + m.join(",") + "]";
                    e = l;
                    return g
                }
                if (h && typeof h == "object") {
                    k = h.length;
                    for (c = 0; c < k; c += 1) {
                        d = h[c];
                        if (typeof d == "string") {
                            g = j(d, n);
                            g && m.push(i(d) + (e ? ": " : ":") + g)
                        }
                    }
                } else
                    for (d in n)
                        if (Object.hasOwnProperty.call(n, d)) {
                            g = j(d, n);
                            g && m.push(i(d) + (e ? ": " : ":") + g)
                        }
                g = m.length === 0 ? "{}" : e ? "{\n" + e + m.join(",\n" + e) + "\n" + l + "}" : "{" + m.join(",") + "}";
                e = l;
                return g
        }
    }

    function i(a) {
        d.lastIndex = 0;
        return d.test(a) ? '"' + a.replace(d, function(a) {
            var b = g[a];
            return typeof b == "string" ? b : "\\u" + ("0000" + a.charCodeAt(0).toString(16)).slice(-4)
        }) + '"' : '"' + a + '"'
    }

    function b(a) {
        return a < 10 ? "0" + a : a
    }
    if (typeof Date.prototype.toJSON != "function") {
        Date.prototype.toJSON = function(a) {
            return isFinite(this.valueOf()) ? this.getUTCFullYear() + "-" + b(this.getUTCMonth() + 1) + "-" + b(this.getUTCDate()) + "T" + b(this.getUTCHours()) + ":" + b(this.getUTCMinutes()) + ":" + b(this.getUTCSeconds()) + "Z" : null
        };
        String.prototype.toJSON = Number.prototype.toJSON = Boolean.prototype.toJSON = function(a) {
            return this.valueOf()
        }
    }
    var c = /[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,
        d = /[\\\"\x00-\x1f\x7f-\x9f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,
        e, f, g = {
            "\b": "\\b",
            "\t": "\\t",
            "\n": "\\n",
            "\f": "\\f",
            "\r": "\\r",
            '"': '\\"',
            "\\": "\\\\"
        },
        h;

    var d;
    e = "";
    f = '\t';
    h = null;
    if (!b || typeof b == "function" || typeof b == "object" && typeof b.length == "number") {
        return j("", {
            "": a
        });
    }
    throw new Error("JSON.stringify")
}

/*
 * json文件导出
 */
function exportJsonFile(excelObj, outputPath, cb) {
    var fileName = excelObj.outname;
    if (fileName[0] == '~') {
        cb(true);
        return;
    }

    fileName = fileName.substring(0, 1).toUpperCase() + fileName.substring(1);

    var exportJson = {};
    var annotations = {};

    try {

        for (var i in excelObj.files) {
            var file = excelObj.files[i];

            // 解析 excel文件
            var excel = xlsx.parse(file);

            if (excel.length <= 0) {
                log4js.error('worksheets.length <= 0 [ %s ][ %d ]', file, excel.length);
                cb(true);
                return;
            };
            var worksheets = excel[0];
            if (worksheets.data.length <= 1) {
                log4js.error('worksheets.data.length <= 1 [ %s ][ %d ]', file, worksheets.data.length);
                cb(true);
                return;
            };
            // 遍历获取表头
            var header = [];
            var data = worksheets.data[0];
            if (data.length <= 1) {
                log4js.error('header data.length <= 1 [ %s ][ %d ]', file, data.length);
                cb(true);
                return;
            };
            for (var i in data) {
                if (data[i]) {
                    header.push(data[i]);
                }
            }
            var annotation = {};
            // 遍历获取注释
            {
                var note = worksheets.data[1];
                if (note.length <= 1) {
                    log4js.error('body note.length <= 1 [ %s ][ %d ]', file, i);
                    cb(true);
                    return;
                };
                // 默认第一个字段为key
                if (!note[0]) {
                    log4js.error("annotation is NULL");
                    cb(true);
                    return;
                }
                for (var n = 0; n < header.length; ++n) {
                    if (note[n]) {
                        annotation[header[n]] = note[n];
                    } else {
                        annotation[header[n]] = '';
                    }

                }
            }
            var baseName = path.basename(file, '.xlsx').split('/').pop();
            baseName = baseName.substring(0, 1).toUpperCase() + baseName.substring(1);

            annotations[baseName] = annotation;

            // 遍历获取数据文件
            for (var i = 2; i < worksheets.data.length; ++i) {
                var data = worksheets.data[i];
                if (data.length <= 1) {
                    log4js.error(worksheets);
                    log4js.error('body data.length <= 1 [ %s ][ %d ]', file, i);
                    cb(true);
                    return;
                };
                // 默认第一个字段为key
                if (!data[0]) {
                    continue;
                }
                var key = data[0];
                var json = {};
                for (var n = 1; n < header.length; ++n) {
                    if (data[n] != null) {
                        if (typeof data[n] === "number") {
                            data[n] = parseFloat(data[n].toFixed(2)); //;
                        }

                        if (data[n][0] === '[' && data[n].slice(-1) === ']') {
                            data[n] = JSON.parse(data[n]);
                        }

                        json[header[n]] = data[n];
                    } else {
                        json[header[n]] = '';
                    }

                }
                exportJson['' + key] = json;
            }
        }



        /*// 版本号
        var time = new Date();
        exportJson['-1'] = {
            'time': util.format('%d-%d-%d[%d:%d:%d]', time.getFullYear(),
                time.getMonth() + 1,
                time.getDate(),
                time.getHours(),
                time.getMinutes(),
                time.getSeconds()),
            'version': '1.0.0'
        }*/

        mkdirSync(outputPath);
        fs.writeFile(outputPath + 'Static' + fileName + '.json', jsonToStr(exportJson));
        log4js.info('exportJsonFile[ %s ]', file);
        cb(false, annotations);
    } catch (e) {
        log4js.error(file);
        log4js.error(e);
        cb(true);
    };
};

/*
 * json文件导出
 */
function exportJsonFileToArray(excelObj, outputPath, cb) {

    var file = excelObj.files[0];

    var fileName = excelObj.outname;
    if (fileName[0] == '~') {
        cb(true);
        return;
    }
    fileName = fileName.substring(0, 1).toUpperCase() + fileName.substring(1);
    var exportJson = {};

    try {
        // 解析 excel文件
        var excel = xlsx.parse(file);
        if (excel.length <= 0) {
            log4js.error('worksheets.length <= 0 [ %s ][ %d ]', file, excel.length);
            cb(true);
            return;
        };
        var worksheets = excel[0];
        if (worksheets.data.length <= 1) {
            log4js.error('worksheets.data.length <= 1 [ %s ][ %d ]', file, worksheets.data.length);
            cb(true);
            return;
        };
        // 遍历获取表头
        var header = [];
        var data = worksheets.data[0];
        if (data.length <= 1) {
            log4js.error('header data.length <= 1 [ %s ][ %d ]', file, data.length);
            cb(true);
            return;
        };
        for (var i in data) {
            if (data[i]) {
                header.push(data[i]);
            }
        }
        var annotation = {};
        // 遍历获取注释
        {
            var note = worksheets.data[1];
            if (note.length <= 1) {
                log4js.error('body note.length <= 1 [ %s ][ %d ]', file, i);
                cb(true);
                return;
            };
            // 默认第一个字段为key
            if (!note[0]) {
                log4js.error("annotation is NULL");
                cb(true);
                return;
            }
            for (var n = 0; n < header.length; ++n) {
                if (note[n]) {
                    annotation[header[n]] = note[n];
                } else {
                    annotation[header[n]] = '';
                }
                exportJson[header[n]] = [];
            }
        }
        // 遍历获取数据文件
        for (var i = 2; i < worksheets.data.length; ++i) {
            var data = worksheets.data[i];
            if (data.length <= 1) {
                log4js.error('body data.length <= 1 [ %s ][ %d ]', file, i);
                cb(true);
                return;
            };

            for (var n = 0; n < header.length; ++n) {
                if (data[n]) {
                    exportJson[header[n]].push(data[n]);
                }
            }
        }
        /*// 版本号
        var time = new Date();
        exportJson['-1'] = {
            'time': util.format('%d-%d-%d[%d:%d:%d]', time.getFullYear(),
                time.getMonth() + 1,
                time.getDate(),
                time.getHours(),
                time.getMinutes(),
                time.getSeconds()),
            'version': '1.0.0'
        }*/

        mkdirSync(outputPath);
        fs.writeFile(outputPath + 'Static' + fileName + '.json', jsonToStr(exportJson));
        log4js.info('exportJsonFile[ %s ]', file);
        cb(false, fileName, annotation);
    } catch (e) {
        log4js.error(file);
        log4js.error(e);
        cb(true);
    };
};

/*
 * 查找文件
 */
function findFiles(inputPath, callback) {
    var exportList = [];
    async.waterfall([
        function(cb) {
            fs.readdir(inputPath, function(err, filesList) {
                cb(err, filesList);
            });
        },
        function(files, cb) {
            var count = 0;
            var len = files.length;

            async.whilst(
                function() {
                    return count < len
                },
                function(cb1) {

                    var file = inputPath + '/' + files[count];
                    count++;

                    if (path.extname(file) == '.xlsx') {

                        exportList.push({
                            'outname': path.basename(file, '.xlsx'),
                            'files': [file]
                        });

                        cb1();

                    } else if (path.extname(file) == '') {

                        fs.readdir(file, function(err, exceles) {

                            if (exceles == null) {
                                cb1();
                                return;
                            }

                            var excelList = [];

                            for (var i in exceles) {

                                var excel = file + '/' + exceles[i];
                                if (exceles[i][0] == '~') {
                                    continue;
                                }

                                if (path.extname(excel) == '.xlsx') {
                                    excelList.push(excel);
                                }
                            }

                            exportList.push({
                                'outname': path.basename(file, '.xlsx'),
                                'files': excelList
                            });

                            cb1();
                        });

                    } else {
                        cb1();
                    }
                },
                function(err) {
                    cb();
                }
            );
        }
    ], function(err) {
        callback(null, exportList);
    });
}

/*
 * 生成Json文件
 */
function oldMakeJsonlFile() {
    var inputPath = process.argv[2];
    var outputPath = process.argv[3] + '/';

    findFiles(inputPath, function(err, exportList) {
        var notes = {};
        var len = exportList.length;
        var count = 0;

        async.whilst(
            function() {
                return count < len
            },
            function(cb) {
                var exportObj = exportList[count];

                var fileNameStr = exportObj.outname;

                if (fileNameStr == "nickname") {
                    exportJsonFileToArray(exportObj, outputPath, function(error, fileName, annotation) {
                        if (!error) {
                            notes[fileName] = annotation;
                        };
                        count++;
                        cb();
                    })
                } else {
                    exportJsonFile(exportObj, outputPath, function(error, annotations) {
                        if (!error) {
                            for (var i in annotations) {
                                notes[i] = annotations[i];
                            }
                        };
                        count++;
                        cb();
                    });
                }
            },
            function(err) {
                fs.writeFile(outputPath + '数据表描述.json', jsonToStr(notes));
            }
        );
    });
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function toInt(obj) {
    return parseInt(obj);
}

function toFloat(obj) {
    return parseFloat(obj);
}

function toArray(obj) {
    return JSON.parse(obj);
}

function toString(obj) {
    return '' + obj;
}

var typeTable = {};
typeTable['number'] = toInt;
typeTable['string'] = toString;
typeTable['array'] = toArray;
typeTable['float'] = toFloat;

function parserMaker(inputname, typename, excel) {

    var sheet = excel[0];
    var head = sheet.data[0];

    var cel = -1;

    for (var i in head) {
        var celname = head[i];

        if (celname === inputname) {
            cel = i;
        }
    }

    var parseData = typeTable[typename];

    return function(row) {
        var obj = sheet.data[row][cel];
        if (obj == null) {
            return "";
        }
        return parseData(obj);
    };
}

function ruleMaker(rule, inputPath, outputPath) {

    if (false == fs.existsSync(inputPath + rule.file)) {
        return function() {
            return;
        };
    }

    var excel = xlsx.parse(inputPath + rule.file);

    var fieldMaker = {};

    for (var i in rule.fields) {
        var field = rule.fields[i];

        if (field.typename === 'undefined') {
            field.typename = 'string';
        }

        var parser = parserMaker(field.inputname, field.typename, excel);
        fieldMaker[field.outname] = parser;
    }

    mkdirSync(outputPath);

    return function() {

        var sheet = excel[0];

        var exportJson = {};
        if (rule.filetype === 'key-obj') {
            for (var i = 2; i < sheet.data.length; i++) {

                var key = null;
                var obj = {};
                for (var j in fieldMaker) {
                    var readData = fieldMaker[j];
                    if (key == null) {
                        key = readData(i);
                        continue;
                    }

                    obj[j] = readData(i);
                }
                exportJson[key] = obj;
            }
        }

        if (rule.filetype === 'key-value') {
            if (Object.keys(fieldMaker).length > 2) {
                for (var i in fieldMaker) {
                    var readData = fieldMaker[i];
                    exportJson[i] = readData(2);
                }
            } else {
                for (var i = 2; i < sheet.data.length; i++) {
                    var key = null;
                    var value = null;
                    for (var j in fieldMaker) {
                        var readData = fieldMaker[j];
                        if (key === null) {
                            key = readData(i);
                            continue;
                        }

                        if (value === null) {
                            value = readData(i);
                            break;
                        }
                    }
                    exportJson[key] = value;
                }
            }
        }

        fs.writeFileSync(outputPath + rule.outfile, jsonToStr(exportJson), 'utf8');
        log4js.info('exportJsonFile[ %s ]', rule.outfile);
    }
}

function makeJsonFile() {
    var inputPath = process.argv[2] + '/';
    var outputPath = process.argv[3] + '/';

    var config = {};
    ruleConfigMaker.makeRuleCfg();

    fs.readFile('./config.json', function(err, data) {
        if (data.length == 0) {
            return;
        }

        config = JSON.parse(data);
        for (var i in config) {
            var ruleJson = config[i];

            var rule = ruleMaker(ruleJson, inputPath, outputPath);

            rule();
        }
    });
}

makeJsonFile();
