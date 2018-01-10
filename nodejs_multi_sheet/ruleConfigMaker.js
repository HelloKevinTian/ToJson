/**
 *  Excel to JSON
 *
 */

var xlsx = require('node-xlsx');
var fs = require('fs');
var crypto = require('crypto');
var path = require('path');
var log4js = require('log4js').getLogger('');
var async = require("async");
var util = require("util");

function makeRuleJson(file) {
    var excel = xlsx.parse('./table/' + file);
    if (excel.length <= 0) {
        log4js.error('worksheets.length <= 0 [ %s ][ %d ]', file, excel.length);
        return;
    }

    var obj = {};

    for (var i = 0; i < excel.length; i++) {
        var o = makeOneSheet(file, i, excel[i]);
        if (o.fields.length > 0) {
            obj[o.outfile] = o;
        }
    };

    return obj;
}

function makeOneSheet(file, index, worksheets) {
    var sheetName = worksheets.name;
    if (worksheets.data.length <= 1) {
        log4js.error('worksheets.data.length <= 1 [ %s ][ %d ]', file, worksheets.data.length);
        return;
    }

    var tableHead = worksheets.data[0]; // 字段名 string
    var tableType = worksheets.data[1]; // 字段类型 string   （worksheets.data[2]是字段注释）
    var tableData = worksheets.data[3]; // 字段正式数据

    var md5 = crypto.createHash('md5');
    md5.update(JSON.stringify(tableHead));
    var res = md5.digest('hex');

    var fields = [];
    for (var i in tableHead) {

        var typename = tableType[i];
        if (typename !== 'string' && typename !== 'int' && typename !== 'float') {
            log4js.error('There is a wrong data type(eg: string int float) in the worksheets: ', file, typename);
            return;
        }

        fields.push({
            'inputname': tableHead[i],
            'outname': tableHead[i],
            'typename': typename
        });

    }

    var outfile = path.basename(file, '.xlsx') + '_' + sheetName + '.json';

    return {
        'file': file,
        'outfile': outfile,
        'index': index,
        'md5': res,
        'filetype': 'key-obj',
        'fields': fields
    };
}

function makeRules() {

    var config = {};
    if (true === fs.existsSync('./config.json')) {
        var data = fs.readFileSync('./config.json', 'utf8');
        config = JSON.parse(data);
    }

    var files = fs.readdirSync('./table');
    for (var i in files) {
        var file = files[i];
        if (path.extname(file) === '.xlsx') {
            var jsonObj = makeRuleJson(file);

            for (var k in jsonObj) {
                var json = jsonObj[k];

                if (config[k] != null && config[k].md5 === json.md5) {
                    continue;
                }

                config[k] = json;
            }

        }
    }

    fs.writeFileSync('./config.json', JSON.stringify(config));
}


/**
 * 导出函数列表
 */
module.exports = {
    'makeRuleCfg': makeRules
};