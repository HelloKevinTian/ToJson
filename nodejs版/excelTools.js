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
typeTable['int'] = toInt;
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
            for (var i = 3; i < sheet.data.length; i++) {
                var key = null;
                var obj = {};
                for (var j in fieldMaker) {
                    var readData = fieldMaker[j];
                    if (key == null) {
                        key = readData(i);
                        // continue; //打开后第一个字段只作为key，不作为obj里的字段
                    }

                    obj[j] = readData(i);

                    //去除字符串中的 \r
                    if (typeof obj[j] === 'string') {
                        obj[j] = obj[j].replace(/\r/g, '');
                    }
                }
                exportJson[key] = obj;
            }
        }


        //暂未使用！
        // if (rule.filetype === 'key-value') {
        //     if (Object.keys(fieldMaker).length > 2) {
        //         for (var i in fieldMaker) {
        //             var readData = fieldMaker[i];
        //             exportJson[i] = readData(2);
        //         }
        //     } else {
        //         for (var i = 3; i < sheet.data.length; i++) {
        //             var key = null;
        //             var value = null;
        //             for (var j in fieldMaker) {
        //                 var readData = fieldMaker[j];
        //                 if (key === null) {
        //                     key = readData(i);
        //                     continue;
        //                 }

        //                 if (value === null) {
        //                     value = readData(i);
        //                     break;
        //                 }
        //             }
        //             exportJson[key] = value;
        //         }
        //     }
        // }

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