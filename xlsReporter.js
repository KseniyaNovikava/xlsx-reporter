var fs = require('fs');
var path = require('canonical-path');
var _ = require('lodash');
var xl = require('excel4node');

// Custom reporter
var Reporter = function(options) {

  var wb = new xl.Workbook();
  var ws = wb.addWorksheet('Sheet 1');

  ws.column(2).setWidth(35);

  var _defaultOutputFile = path.resolve(process.cwd(), './_test-output', 'protractor-results.xlsx');
  options.outputFile = options.outputFile || _defaultOutputFile;

  initOutputFile(options.outputFile);
  options.appDir = options.appDir ||  './';
  var _root = { appDir: options.appDir, suites: [] };
  log('AppDir: ' + options.appDir, +1);
  var _currentSuite;

  this.suiteStarted = function(suite) {
    _currentSuite = { description: suite.description, status: null, specs: [] };
    _root.suites.push(_currentSuite);
    log('Suite: ' + suite.description, +1);
  };

  this.suiteDone = function(suite) {
    var statuses = _currentSuite.specs.map(function(spec) {
      return spec.status;
    });
    statuses = _.uniq(statuses);
    var status = statuses.indexOf('failed') >= 0 ? 'failed' : statuses.join(', ');
    _currentSuite.status = status;
    log('Suite ' + _currentSuite.status + ': ' + suite.description, -1);
  };

  this.specStarted = function(spec) {

  };

  this.specDone = function(spec) {
    var currentSpec = {
      description: spec.description,
      status: spec.status
    };
    if (spec.failedExpectations.length > 0) {
      currentSpec.failedExpectations = spec.failedExpectations;
    }

    _currentSuite.specs.push(currentSpec);
    log(spec.status + ' - ' + spec.description);
  };

  this.jasmineDone = function() {
    outputFile = options.outputFile;
    var output = formatOutput(_root);
    wb.write(outputFile);
    //ws.cell(4, 1).string(output);

  };

  function ensureDirectoryExistence(filePath) {
    var dirname = path.dirname(filePath);
    if (directoryExists(dirname)) {
      return true;
    }
    ensureDirectoryExistence(dirname);
    fs.mkdirSync(dirname);
  }

  function directoryExists(path) {
    try {
      return fs.statSync(path).isDirectory();
    }
    catch (err) {
      return false;
    }
  }

  function initOutputFile(outputFile) {
    ensureDirectoryExistence(outputFile);
    var header = "Protractor results for: " + (new Date()).toLocaleString() + "\n\n";
    wb.write(outputFile);
    ws.cell(2, 1, 2, 3, true).string(header);
  }


    function formatOutput(output) {
    var indent = '  ';
    var pad = '  ';
    var results=[];
    var i=4;
    ws.cell(i, 1).string('AppDir:' + output.appDir);
   output.suites.forEach(function(suite) {
     ws.cell(++i, 1).string(pad + 'Suite: ');
      ws.cell(i, 2).string(suite.description);
       ws.cell(i, 3).string(suite.status);
      pad+=indent;
      suite.specs.forEach(function(spec) {
        ws.cell(++i, 1).string(pad + spec.status);
        ws.cell(i, 2).string(spec.description);
        if (spec.failedExpectations) {
          pad+=indent;
          spec.failedExpectations.forEach(function (fe) {
            ws.cell(i, 3).string(pad + 'message: ' + fe.message);
          });
          pad=pad.substr(2);
        }
      });
      pad = pad.substr(2);
    });
   return wb.write(outputFile);
 
  }


  // for console output
  var _pad;
  function log(str, indent) {
    _pad = _pad || '';
    if (indent == -1) {
      _pad = _pad.substr(2);
    }
    console.log(_pad + str);
    if (indent == 1) {
      _pad = _pad + '  ';
    }
  }
};

module.exports = Reporter;
