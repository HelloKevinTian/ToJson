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

    var worksheets = excel[0];
    if (worksheets.data.length <= 1) {
        log4js.error('worksheets.data.length <= 1 [ %s ][ %d ]', file, worksheets.data.length);
        cb(true);
        return;
    }

    var tableHead = worksheets.data[0];
    var tableData = worksheets.data[2];

    var md5 = crypto.createHash('md5');
    md5.update(JSON.stringify(tableHead));
    var res = md5.digest('hex');

    var fields = [];
    for (var i in tableHead) {
        var typename = typeof tableData[i];
        if (typename === 'string') {

            if (tableData[i][0] === '[' && tableData[i].slice(-1) === ']') {
                typename = 'array';
            }
        }

        fields.push({
            'inputname': tableHead[i],
            'outname': tableHead[i],
            'typename': typename
        });
    }

    var outfile = path.basename(file, '.xlsx') + '.json';

    return {
        'file': file,
        'outfile': outfile,
        'md5': res,
        'filetype':'key-obj',
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
            var json = makeRuleJson(file);

            if (config[file] != null && config[file].md5 === json.md5) {
                continue;
            }

            config[file] = json;
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
