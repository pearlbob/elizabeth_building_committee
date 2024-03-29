import 'dart:collection';
import 'dart:io';
import 'dart:math';

import 'package:elizabeth_building_committee/homePath.dart';
import 'package:excel/excel.dart';

const _columnNameValveNumber = 'VAL #';
const _columnNameFunction = ' FUNCTION';

class ElizabethValveChart {
  ElizabethValveChart() {
    _styles[_columnNameValveNumber] = _dataCellRightStyle;
  }

  Future<bool> processFile(String filePathAsString) async {
    await input(filePathAsString);
    _process();
    await output();
    return true;
  }

  Future<bool> input(String filePathAsString) async {
    var bytes = File(filePathAsString).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    _columnNames = [];
    _rows = [];

    for (var sheetName in excel.tables.keys) {
      print(sheetName); //sheet Name
      var sheet = excel.tables[sheetName];
      if (sheet == null) {
        continue;
      }

      var r = 0;
      for (var row in sheet.rows ?? []) {
        if (r == 0) {
          _columnNames = row;
        } else {
          var map = <String, String>{};
          var i = 0;
          for (var columnName in _columnNames) {
            map[columnName.toString()] = row[i].toString();
            i++;
          }
          _rows.add(map);
        }
        r++;
      }
    }

    print('size: (${_rows.length},${_columnNames.length})');

    // for (var i = 0; i < _columnNames.length; i++) {
    //   print(' $i,  // ${_columnNames[i]}');
    // }
    return true;
  }

  final _titleCellStyle = CellStyle(
    //backgroundColorHex: '#1AFF1A',
    fontFamily: getFontFamily(FontFamily.Calibri),
    bold: true,
    textWrapping: TextWrapping.WrapText,
    horizontalAlign: HorizontalAlign.Center,
  );

  final _dataCellRightStyle = CellStyle(
    //fontFamily: getFontFamily(FontFamily.Calibri),
    textWrapping: TextWrapping.WrapText,
    horizontalAlign: HorizontalAlign.Right,
  );
  final _dataCellLeftStyle = CellStyle(
    //fontFamily: getFontFamily(FontFamily.Calibri),
    textWrapping: TextWrapping.WrapText,
    horizontalAlign: HorizontalAlign.Left,
  );
  final _dataCellCenterStyle = CellStyle(
    //fontFamily: getFontFamily(FontFamily.Calibri),
    textWrapping: TextWrapping.WrapText,
    horizontalAlign: HorizontalAlign.Center,
  );
  final _dataCellClipStyle = CellStyle(
    //fontFamily: getFontFamily(FontFamily.Calibri),
    textWrapping: TextWrapping.Clip,
    horizontalAlign: HorizontalAlign.Center,
  );

  void _process() {
    //  sort the rows
    for (var map in _rows) {
      _orderedRows.add(_OrderedRow(map));
    }

    for (var row in _orderedRows) {
      var valveFunction = row.map[_columnNameFunction] ?? '';

      //  most specific first

      //  look for columns
      {
        final _columnAmpersandThruRegexp =
            RegExp(r'(\d\d) *\& *(\d\d) UNITS (\d\d?)TH THRU.? (\d\d?)TH FLOORS', caseSensitive: false);
        var matches = _columnAmpersandThruRegexp.allMatches(valveFunction);
        if (matches.isNotEmpty) {
          for (var m in matches) {
            var firstUnit = int.parse(m.group(1) ?? '0');
            var lastUnit = int.parse(m.group(2) ?? '0');
            var firstFloor = int.parse(m.group(3) ?? '0');
            var lastFloor = int.parse(m.group(4) ?? '0');
            // print(
            //     '_columnAmpersandRegexp: $firstUnit & $lastUnit on floors $firstFloor -> $lastFloor:  $valveFunction');
            assert(firstUnit < lastUnit || (firstUnit == 17 && lastUnit == 1));
            assert(firstFloor < lastFloor);
            assert(firstFloor >= 4);
            assert(lastFloor <= 13);
            for (var floor = firstFloor; floor <= lastFloor; floor++) {
              addByUnitMatches(floor * 100 + firstUnit, row);
              addByUnitMatches(floor * 100 + lastUnit, row);
            }
          }
          continue;
        }
      }

      //  look for THRU units
      {
        final _thruUnitNumberRegexp = RegExp(r'(\d\d\d\d?) +thru\.? +(\d\d\d\d?)', caseSensitive: false);
        var matches = _thruUnitNumberRegexp.allMatches(valveFunction);
        if (matches.isNotEmpty) {
          for (var m in matches) {
            var firstUnit = int.parse(m.group(1) ?? '0');
            var lastUnit = int.parse(m.group(2) ?? '0');
            assert(firstUnit < lastUnit);
            assert(firstUnit % 100 == lastUnit % 100);
            //print('thru ${m.group(0)}: $firstUnit -> $lastUnit');
            for (var unit = firstUnit; unit <= lastUnit; unit += 100) {
              addByUnitMatches(unit, row);
            }
          }
          continue;
        }
      }

      //  look for - units
      {
        final _dashUnitNumberRegexp = RegExp(r'(\d\d\d\d?) *- *(\d\d\d\d?)');
        var matches = _dashUnitNumberRegexp.allMatches(valveFunction);
        if (matches.isNotEmpty) {
          for (var m in matches) {
            var firstUnit = int.parse(m.group(1) ?? '0');
            var lastUnit = int.parse(m.group(2) ?? '0');
            //print('dash ${m.group(0)}: $firstUnit -> $lastUnit');
            assert(firstUnit < lastUnit);
            assert(firstUnit % 100 == lastUnit % 100);

            for (var unit = firstUnit; unit <= lastUnit; unit += 100) {
              addByUnitMatches(unit, row);
            }
          }
          continue;
        }
      }

      {
        final _columnAmpersandItemsThruRegexp = RegExp(
            r'(\d\d) *\& *(\d\d) (?:kitchens|baths|units|BATH/KITCHEN|KITCHEN/BATH),?'
            r' (\d\d?)(?:rd|th)-(\d\d?)th ?(?:FL)?',
            caseSensitive: false);
        var matches = _columnAmpersandItemsThruRegexp.allMatches(valveFunction);
        if (matches.isNotEmpty) {
          for (var m in matches) {
            var firstUnit = int.parse(m.group(1) ?? '0');
            var lastUnit = int.parse(m.group(2) ?? '0');
            var firstFloor = int.parse(m.group(3) ?? '0');
            var lastFloor = int.parse(m.group(4) ?? '0');
            assert(firstUnit < lastUnit || (firstUnit == 17 && lastUnit == 1));
            assert(firstFloor < lastFloor);
            assert(firstFloor >= 3);
            assert(lastFloor <= 13);
            //      print('_columnAmpersandItemsThruRegexp: $firstUnit -> $lastUnit, $firstFloor-${lastFloor}FL : $valveFunction');
            for (var floor = min(4, firstFloor); floor <= lastFloor; floor++) {
              addByUnitMatches(floor * 100 + firstUnit, row);
              addByUnitMatches(floor * 100 + lastUnit, row);
            }
          }
          continue;
        }
      }

      {
        final _columnUnitItemsThruRegexp = RegExp(
            r'(\d\d) * (?:kitchens|baths|units|BATH/KITCHEN),? (\d\d?)(?:rd|th) +thru +(\d\d?)th FL',
            caseSensitive: false);
        var matches = _columnUnitItemsThruRegexp.allMatches(valveFunction);
        if (matches.isNotEmpty) {
          for (var m in matches) {
            var firstUnit = int.parse(m.group(1) ?? '0');
            var firstFloor = int.parse(m.group(2) ?? '0');
            var lastFloor = int.parse(m.group(3) ?? '0');
            assert(firstFloor < lastFloor);
            assert(firstFloor >= 3);
            assert(lastFloor <= 13);
            // print('     _columnUnitItemsThruRegexp: $firstUnit, ${firstFloor}-${lastFloor}FL : $valveFunction');
            for (var floor = min(4, firstFloor); floor <= lastFloor; floor++) {
              addByUnitMatches(floor * 100 + firstUnit, row);
            }
          }
          continue;
        }
      }

      {
        final _columnItemsThruRegexp = RegExp(
            r' (\d\d) (?:kitchens|baths|units|BATH/KITCHEN) (\d\d?)(?:rd|th)-(\d\d?)th FL',
            caseSensitive: false);
        var matches = _columnItemsThruRegexp.allMatches(valveFunction);
        if (matches.isNotEmpty) {
          for (var m in matches) {
            var firstUnit = int.parse(m.group(1) ?? '0');
            var firstFloor = int.parse(m.group(2) ?? '0');
            var lastFloor = int.parse(m.group(3) ?? '0');
            assert(firstFloor < lastFloor);
            assert(firstFloor >= 3);
            assert(lastFloor <= 13);
            // print('        _columnItemsThruRegexp: $firstUnit, ${firstFloor}-${lastFloor}FL : $valveFunction');
            for (var floor = min(4, firstFloor); floor <= lastFloor; floor++) {
              addByUnitMatches(floor * 100 + firstUnit, row);
            }
          }
          continue;
        }
      }
      {
        final _columnComboItemsThruRegexp = RegExp(
            r' (\d\d) (?:kitchens|baths|units|BATH/KITCHEN|KITCHEN/BATH), (\d\d) (?:kitchens|bath|baths|units|BATH/KITCHEN|KITCHEN/BATH)'
            r' (\d\d?)(?:rd|th)-(\d\d?)th FL',
            caseSensitive: false);
        var matches = _columnComboItemsThruRegexp.allMatches(valveFunction);
        if (matches.isNotEmpty) {
          for (var m in matches) {
            var firstUnit = int.parse(m.group(1) ?? '0');
            var lastUnit = int.parse(m.group(2) ?? '0');
            var firstFloor = int.parse(m.group(3) ?? '0');
            var lastFloor = int.parse(m.group(4) ?? '0');
            assert(firstFloor < lastFloor);
            assert(firstFloor >= 4);
            assert(lastFloor <= 13);
            // print('            _columnComboItemsThruRegexp:  $firstUnit-$lastUnit on $firstFloor-$lastFloor floor, $valveFunction');
            for (var floor = firstFloor; floor <= lastFloor; floor++) {
              addByUnitMatches(floor * 100 + firstUnit, row);
              addByUnitMatches(floor * 100 + lastUnit, row);
            }
          }
          continue;
        }
      }

      {
        // CIRCUIT SETTER 05 & 06 KITCHENS 8TH-13TH FL
        // CIRCUIT SETTER 08 KITCHEN/BATH & 09 KITCHEN 8TH-13TH
        final _circuitSetterKitchensBathsRegexp = RegExp(
            r'^ *CIRCUIT SETTER +(\d\d) (?:kitchens|kitchen/bath)? *\& +(\d\d) *(?:kitchens|baths|kitchen)'
            r' (\d\d?)(?:rd|th)-(\d\d?)th *(?:fl)? *$',
            caseSensitive: false);
        var matches = _circuitSetterKitchensBathsRegexp.allMatches(valveFunction);
        if (matches.isNotEmpty) {
          for (var m in matches) {
            var firstUnit = int.parse(m.group(1) ?? '0');
            var lastUnit = int.parse(m.group(2) ?? '0');
            var firstFloor = int.parse(m.group(3) ?? '0');
            var lastFloor = int.parse(m.group(4) ?? '0');
            print('            _circuitSetterKitchensBathsRegexp: v${row.valveNumber}:'
                ' $firstUnit-$lastUnit on $firstFloor-$lastFloor floor, $valveFunction');
            for (var floor = firstFloor; floor <= lastFloor; floor++) {
              addByUnitMatches(floor * 100 + firstUnit, row);
              addByUnitMatches(floor * 100 + lastUnit, row);
            }
          }
          continue;
        }
      }

      {
        var firstFloor = 4;
        var lastFloor = 13;
        // CIRCUIT SETTER 09 & 10 BATHS
        // CIRCUIT SETTER 10 & 11 KITCHENS
        // CIRCUIT SETTER 10 KITCHENS & 11 BATHS
        // CIRCUIT SETTER 09 KITCHENS & 08 BATHS
        final _circuitSetterKitchensBathsRegexp = RegExp(
            r'^ *CIRCUIT SETTER +(\d\d) (?:kitchens|kitchen/bath)? *\& +(\d\d) *(?:kitchens|baths|kitchen)'
            r' *$',
            caseSensitive: false);
        var matches = _circuitSetterKitchensBathsRegexp.allMatches(valveFunction);
        if (matches.isNotEmpty) {
          for (var m in matches) {
            var firstUnit = int.parse(m.group(1) ?? '0');
            var lastUnit = int.parse(m.group(2) ?? '0');
            // print('            _circuitSetterKitchensBathsRegexp: v${row.valveNumber}:'
            //     ' $firstUnit-$lastUnit on $firstFloor-$lastFloor floor, $valveFunction');
            for (var floor = firstFloor; floor <= lastFloor; floor++) {
              addByUnitMatches(floor * 100 + firstUnit, row);
              addByUnitMatches(floor * 100 + lastUnit, row);
            }
          }
          continue;
        }
      }

      {
        var firstFloor = 4;
        var lastFloor = 13;
        // CIRCUIT SETTER 08 KITCHENS
        final _circuitSetterKitchensRegexp = RegExp(r'^ *CIRCUIT SETTER +(\d\d) (?:kitchens) *$', caseSensitive: false);
        var matches = _circuitSetterKitchensRegexp.allMatches(valveFunction);
        if (matches.isNotEmpty) {
          for (var m in matches) {
            var unit = int.parse(m.group(1) ?? '0');
            print(
                '            _circuitSetterKitchensRegexp: v${row.valveNumber}: $unit on $firstFloor-$lastFloor floor, $valveFunction');
            for (var floor = firstFloor; floor <= lastFloor; floor++) {
              addByUnitMatches(floor * 100 + unit, row);
            }
          }
          continue;
        }
      }

      // {
      //   final _forFloorRegexp = RegExp(r'(\d\d?)TH FLoor', caseSensitive: false);
      //   var matches = _forFloorRegexp.allMatches(valveFunction);
      //   if (matches.isNotEmpty) {
      //     for (var m in matches) {
      //       var firstFloor = int.parse(m.group(1) ?? '0');
      //       print('      _forFloorRegexp: ${row.valveNumber}: $firstFloor,  $valveFunction');
      //     }
      //     continue;
      //   }
      // }
      //
      // {
      //   final _forFlRegexp = RegExp(r'(\d\d?)TH FL', caseSensitive: false);
      //   var matches = _forFlRegexp.allMatches(valveFunction);
      //   if (matches.isNotEmpty) {
      //     for (var m in matches) {
      //       print('_forFlRegexp: ${row.valveNumber}: ${m.group(0)} $valveFunction');
      //     }
      //     continue;
      //   }
      // }

      {
        //  look for individually named units
        final _unitNumberRegexp = RegExp(r'(\d\d\d\d?)');
        for (var m in _unitNumberRegexp.allMatches(valveFunction)) {
          var unit = int.parse(m.group(1) ?? '0');
          addByUnitMatches(unit, row);
        }
      }
    }

    //  find all the valves used in byUnit
    {
      var units = SplayTreeSet<int>();
      var valves = SplayTreeSet<int>();
      for (var unit in byUnit.keys) {
        units.add(unit);
        for (var orderedRow in byUnit[unit]?.toList() ?? []) {
          valves.add(orderedRow.valveNumber);
        }
      }

      for (var row in _orderedRows) {
        if (valves.contains(row.valveNumber)) {
          continue;
        }
        //print('$row');
        addByUnitMatches(9999, row); //  valves unassociated with a unit
      }

      if (false) {
        var lastUnit = 0;
        for (var unit in units) {
          if (unit != lastUnit + 1) {
            if (lastUnit > 0) {
              print('   to unit $lastUnit');
            }
            print('from unit $unit');
          }
          lastUnit = unit;
        }
        print('   to unit ${units.last}');

        print('');
        print('valves');
        var lastValve = 0;
        for (var valve in valves) {
          if (valve != lastValve + 1) {
            if (lastValve > 0) {
              print('   to valve $lastValve');
            }
            print('from valve $valve');
          }
          lastValve = valve;
        }
        print('   to valve ${valves.last}');
      }
    }
    // for (var unit in SplayTreeSet.from(byUnit.keys)) {
    //   print('$unit:');
    //   for (var row in byUnit[unit] ?? <_OrderedRow>[]) {
    //     print('     $row');
    //   }
    // }
  }

  void addByUnitMatches(int unit, _OrderedRow row) {
    var unitOrderedRows = byUnit[unit];
    if (unitOrderedRows == null) {
      byUnit[unit] = SplayTreeSet();
    }
    byUnit[unit]!.add(row);
  }

  Future<bool> output() async {
    var excel = Excel.createExcel();

    var outputColumns = <int>[
      0, //
      1, //
      2, //
      3, //
      4, //
      5, //
    ];

    {
      excel.rename(await excel.getDefaultSheet(), 'by unit');
      var sheet = excel.sheets.values.first;

      {
        sheet.updateCell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: 0), 'Unit', cellStyle: _titleCellStyle);
        var i = 0;
        for (var outputColumn in outputColumns) {
          sheet.updateCell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: i + 1), _columnNames[outputColumn],
              cellStyle: _titleCellStyle);
          i++;
        }
      }

      {
        var r = 0;
        var lastUnit = 0;
        for (var unit in SplayTreeSet.from(byUnit.keys)) {
          for (var row in byUnit[unit] ?? <_OrderedRow>[]) {
            sheet.updateCell(
                CellIndex.indexByColumnRow(rowIndex: r + 1, columnIndex: 0), (unit != lastUnit ? unit.toString() : ''),
                cellStyle: _dataCellRightStyle);
            lastUnit = unit;
            var i = 1;
            for (var outputColumn in outputColumns) {
              var name = _columnNames[outputColumn];
              sheet.updateCell(CellIndex.indexByColumnRow(rowIndex: r + 1, columnIndex: i++), row.value(name),
                  cellStyle: _styles[name] ?? _dataCellLeftStyle);
            }
            r++;
          }
        }
      }

      print('rows: ${sheet.rows.length}');
      print('maxCols: ${sheet.maxCols}');
    }

    // for (var key in excel.sheets.keys) {
    //   print('key: $key');
    //   var sheet = excel.sheets[key];
    //   print('   sheet: ${sheet?.sheetName}');
    //   for (var row in sheet?.rows ?? []) {
    //     print('      row: $row');
    //   }
    // }

    await excel.encode().then((onValue) {
      File('${homePath()}/junk/excel_written.xlsx')
        ..createSync(recursive: true)
        ..writeAsBytesSync(onValue);
    });

    return true;
  }

  void diagnosticPrint() {
    var r = 0;
    for (var map in _rows) {
      print('${r++}:');
      for (var key in _columnNames) {
        print('   $key: ${map[key]}');
      }
    }
  }

  var _columnNames = [];
  final _styles = <String, CellStyle>{};
  var _rows = <Map<String, String>>[];
  final _orderedRows = SplayTreeSet<_OrderedRow>();

  Map<int, Set<_OrderedRow>> byUnit = {};
}

class _OrderedRow implements Comparable<_OrderedRow> {
  _OrderedRow(this.map) : valveNumber = int.parse(map[_columnNameValveNumber].toString());

  dynamic value(String name) {
    switch (name) {
      case _columnNameValveNumber:
        return valveNumber;
      default:
        return map[name] ?? '';
    }
  }

  int valveNumber;

  var map = <String, String>{};

  @override
  int compareTo(_OrderedRow other) {
    int? ret = valveNumber.compareTo(other.valveNumber);
    if (ret != 0) {
      return ret;
    }

    for (var id in [
      _columnNameValveNumber,
    ]) {
      ret = map[id]?.compareTo(other.map[id] ?? '');
      if (ret != null && ret != 0) {
        return ret;
      }
    }

    return 0;
  }

  @override
  String toString() {
    return '{valve: $valveNumber,'
        ' function: ${map[_columnNameFunction]}'
        //' map: $map'
        '}';
  }
}

int _unitMax(int floor) {
  switch (floor) {
    case 1:
      return 15; //  ignoring room 141
    case 14:
      return 7;
    case 15:
      return 5;
    default:
      return 17;
  }
}
