"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports["default"] = parseXlsx;
var _fs = _interopRequireDefault(require("fs"));
var _stream = _interopRequireDefault(require("stream"));
var _unzipper = _interopRequireDefault(require("unzipper"));
var _xpath = _interopRequireDefault(require("xpath"));
var _xmldom = _interopRequireDefault(require("@xmldom/xmldom"));
function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { "default": obj }; }
function _typeof(obj) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (obj) { return typeof obj; } : function (obj) { return obj && "function" == typeof Symbol && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; }, _typeof(obj); }
function _createForOfIteratorHelper(o, allowArrayLike) { var it = typeof Symbol !== "undefined" && o[Symbol.iterator] || o["@@iterator"]; if (!it) { if (Array.isArray(o) || (it = _unsupportedIterableToArray(o)) || allowArrayLike && o && typeof o.length === "number") { if (it) o = it; var i = 0; var F = function F() {}; return { s: F, n: function n() { if (i >= o.length) return { done: true }; return { done: false, value: o[i++] }; }, e: function e(_e) { throw _e; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var normalCompletion = true, didErr = false, err; return { s: function s() { it = it.call(o); }, n: function n() { var step = it.next(); normalCompletion = step.done; return step; }, e: function e(_e2) { didErr = true; err = _e2; }, f: function f() { try { if (!normalCompletion && it["return"] != null) it["return"](); } finally { if (didErr) throw err; } } }; }
function _unsupportedIterableToArray(o, minLen) { if (!o) return; if (typeof o === "string") return _arrayLikeToArray(o, minLen); var n = Object.prototype.toString.call(o).slice(8, -1); if (n === "Object" && o.constructor) n = o.constructor.name; if (n === "Map" || n === "Set") return Array.from(o); if (n === "Arguments" || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)) return _arrayLikeToArray(o, minLen); }
function _arrayLikeToArray(arr, len) { if (len == null || len > arr.length) len = arr.length; for (var i = 0, arr2 = new Array(len); i < len; i++) arr2[i] = arr[i]; return arr2; }
function _defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, _toPropertyKey(descriptor.key), descriptor); } }
function _createClass(Constructor, protoProps, staticProps) { if (protoProps) _defineProperties(Constructor.prototype, protoProps); if (staticProps) _defineProperties(Constructor, staticProps); Object.defineProperty(Constructor, "prototype", { writable: false }); return Constructor; }
function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }
function _defineProperty(obj, key, value) { key = _toPropertyKey(key); if (key in obj) { Object.defineProperty(obj, key, { value: value, enumerable: true, configurable: true, writable: true }); } else { obj[key] = value; } return obj; }
function _toPropertyKey(arg) { var key = _toPrimitive(arg, "string"); return _typeof(key) === "symbol" ? key : String(key); }
function _toPrimitive(input, hint) { if (_typeof(input) !== "object" || input === null) return input; var prim = input[Symbol.toPrimitive]; if (prim !== undefined) { var res = prim.call(input, hint || "default"); if (_typeof(res) !== "object") return res; throw new TypeError("@@toPrimitive must return a primitive value."); } return (hint === "string" ? String : Number)(input); }
var ns = {
  a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
};
var select = _xpath["default"].useNamespaces(ns);
function extractFiles(path, sheet) {
  var files = _defineProperty({
    strings: {},
    sheet: {},
    'xl/sharedStrings.xml': 'strings'
  }, "xl/worksheets/sheet".concat(sheet, ".xml"), 'sheet');
  var stream = path instanceof _stream["default"] ? path : _fs["default"].createReadStream(path);
  return new Promise(function (resolve, reject) {
    var filePromises = [];
    stream.pipe(_unzipper["default"].Parse()).on('error', reject).on('close', function () {
      Promise.all(filePromises).then(function () {
        return resolve(files);
      });
    })
    // For some reason `end` event is not emitted.
    // .on('end', () => {
    //   Promise.all(filePromises).then(() => resolve(files));
    // })
    .on('entry', function (entry) {
      var file = files[entry.path];
      if (file) {
        var contents = '';
        filePromises.push(new Promise(function (resolve) {
          entry.on('data', function (data) {
            return contents += data.toString();
          }).on('end', function () {
            files[file].contents = contents;
            resolve();
          });
        }));
      } else {
        entry.autodrain();
      }
    });
  });
}
function calculateDimensions(cells) {
  var comparator = function comparator(a, b) {
    return a - b;
  };
  var allRows = cells.map(function (cell) {
    return cell.row;
  }).sort(comparator);
  var allCols = cells.map(function (cell) {
    return cell.column;
  }).sort(comparator);
  var minRow = allRows[0];
  var maxRow = allRows[allRows.length - 1];
  var minCol = allCols[0];
  var maxCol = allCols[allCols.length - 1];
  return [{
    row: minRow,
    column: minCol
  }, {
    row: maxRow,
    column: maxCol
  }];
}
function extractData(files) {
  var sheet;
  var values;
  var data = [];
  try {
    sheet = new _xmldom["default"].DOMParser().parseFromString(files.sheet.contents);
    var valuesDoc = new _xmldom["default"].DOMParser().parseFromString(files.strings.contents);
    values = select('//a:si', valuesDoc).map(function (string) {
      return select('.//a:t[not(ancestor::a:rPh)]', string).map(function (t) {
        return t.textContent;
      }).join('');
    });
  } catch (parseError) {
    return [];
  }
  function colToInt(col) {
    var letters = ["", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
    col = col.trim().split('');
    var n = 0;
    for (var i = 0; i < col.length; i++) {
      n *= 26;
      n += letters.indexOf(col[i]);
    }
    return n;
  }
  ;
  var na = {
    textContent: ''
  };
  var CellCoords = /*#__PURE__*/_createClass(function CellCoords(cell) {
    _classCallCheck(this, CellCoords);
    cell = cell.split(/([0-9]+)/);
    this.row = parseInt(cell[1]);
    this.column = colToInt(cell[0]);
  });
  var Cell = /*#__PURE__*/_createClass(function Cell(cellNode) {
    _classCallCheck(this, Cell);
    var r = cellNode.getAttribute('r');
    var type = cellNode.getAttribute('t') || '';
    var value = (select('a:v', cellNode, 1) || na).textContent;
    var coords = new CellCoords(r);
    this.column = coords.column;
    this.row = coords.row;
    this.value = value;
    this.type = type;
  });
  var cells = select('/a:worksheet/a:sheetData/a:row/a:c', sheet).map(function (node) {
    return new Cell(node);
  });
  var d = select('//a:dimension/@ref', sheet, 1);
  if (d) {
    d = d.textContent.split(':').map(function (_) {
      return new CellCoords(_);
    });
  } else {
    d = calculateDimensions(cells);
  }
  var cols = d[1].column - d[0].column + 1;
  var rows = d[1].row - d[0].row + 1;
  times(rows, function () {
    var row = [];
    times(cols, function () {
      return row.push('');
    });
    data.push(row);
  });
  var _iterator = _createForOfIteratorHelper(cells),
    _step;
  try {
    for (_iterator.s(); !(_step = _iterator.n()).done;) {
      var cell = _step.value;
      var value = cell.value;
      if (cell.type == 's') {
        value = values[parseInt(value)];
      }
      if (data[cell.row - d[0].row]) {
        data[cell.row - d[0].row][cell.column - d[0].column] = value;
      }
    }
  } catch (err) {
    _iterator.e(err);
  } finally {
    _iterator.f();
  }
  return data;
}
function parseXlsx(path) {
  var sheet = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : '1';
  return extractFiles(path, sheet).then(function (files) {
    return extractData(files);
  });
}
;
function times(n, action) {
  var i = 0;
  while (i < n) {
    action();
    i++;
  }
}
//# sourceMappingURL=excelParser.js.map