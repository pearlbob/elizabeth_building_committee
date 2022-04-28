import 'dart:collection';
import 'dart:io';

import 'package:csv/csv.dart';
import 'package:elizabeth_building_committee/committe_report_entry.dart';
import 'package:elizabeth_building_committee/eliz_5_year_reserve_plan.dart';
import 'package:elizabeth_building_committee/elizabethValveChart.dart';
import 'package:elizabeth_building_committee/homePath.dart';
import 'package:excel/excel.dart';

var entries = <CommitteeReportEntry>[];

void main(List<String> arguments) async {
  //await testCSV();

  //await testXlsx();

  // await ElizabethBuildingCommittee()
  //     .processFile('${homePath()}/junk/excel_file.xlsx');
  //  this doesn't work!!!!   await ElizabethValveChart().processFile('${homePath()}/junk/excel_valve_chart.xlsx');

  Elizabeth5YearReservePlan(File('Eliz_5_year_reserve_plan.csv'));

  exit(0);
}

enum _State {
  initial,
  vendor,
  vendorData,
}

Future<bool> testCSV() async {
  var csvAsString = File('./lib/assets/maintenance_test_report_20210128.csv').readAsStringSync();
  //print( csvAsString);
  var rowsAsListOfValues = const CsvToListConverter().convert(csvAsString);
  var state = _State.initial;
  for (var row in rowsAsListOfValues) {
    if (row.length > 2) {
      switch (state) {
        case _State.vendor:
          //  shouldn't happen
          state = _State.initial;
          break;
        case _State.initial:
          var vendor = row[1];
          if (vendor.compareTo('Vendor') == 0) {
            state = _State.vendorData;
          }
          break;
        case _State.vendorData:
          //  Aug-20,Vendor,Full Focus WO Reference,Task/Item Description,Type,Budgeted Item,Reserve ,Cost ,Status
          String vendor = row[1];

          if (vendor.compareTo('Vendor') == 0) {
            state = _State.vendorData;
            break;
          }
          var entry = CommitteeReportEntry(
            CommitteeReportEntrySource.FAM_report_csv,
            description: row[3],
            dateString: row[0],
            vendor: vendor,
            fullFocusWOReference: row[2],
            type: row[4],
            budgetedItem: row[5],
            reserve: row[6],
            cost: row[7],
            status: row[8],
          );

          if (entry.dateString.isEmpty && entry.vendor.isEmpty && entry.description.isEmpty) {
            state = _State.initial;
            break;
          }

          entries.add(entry);
          break;
      }
    } else {
      state = _State.initial;
    }
  }

  return true;
}

Future<bool> testXlsx() async {
  var fileName = '${homePath()}/junk/excel_file.xlsx';
  var bytes = File(fileName).readAsBytesSync();
  var excel = Excel.decodeBytes(bytes);

  var columnNames = [];
  var rows = <Map<String, String>>{};

  for (var sheetName in excel.tables.keys) {
    print(sheetName); //sheet Name
    var sheet = excel.tables[sheetName];
    if (sheet == null) {
      continue;
    }
    print('size: (${sheet.maxRows},${sheet.maxCols})');

    var r = 0;
    for (var row in sheet.rows ?? []) {
      if (r == 0) {
        columnNames = row;
      } else {
        var map = <String, String>{};
        var i = 0;
        for (var columnName in columnNames) {
          map[columnName.toString()] = row[i].toString();
          i++;
        }
        rows.add(map);
      }
      r++;
    }

    //  diagnostic
    {
      var r = 0;
      for (var map in rows) {
        print('${r++}:');
        for (var key in columnNames) {
          print('   $key: ${map[key]}');
        }
      }
    }
  }

  var titleCellStyle = CellStyle(
    //backgroundColorHex: '#1AFF1A',
    fontFamily: getFontFamily(FontFamily.Calibri),
    bold: true,
    textWrapping: TextWrapping.WrapText,
    // underline: Underline.Single,
  );

  var dataCellRightStyle = CellStyle(
    //fontFamily: getFontFamily(FontFamily.Calibri),
    //textWrapping: TextWrapping.WrapText,
    horizontalAlign: HorizontalAlign.Right,
  );
  var dataCellLeftStyle = CellStyle(
    //fontFamily: getFontFamily(FontFamily.Calibri),
    textWrapping: TextWrapping.WrapText,
    horizontalAlign: HorizontalAlign.Left,
  );
  var dataCellCenterStyle = CellStyle(
    //fontFamily: getFontFamily(FontFamily.Calibri),
    textWrapping: TextWrapping.WrapText,
    horizontalAlign: HorizontalAlign.Center,
  );

  var sheet = excel.tables[excel.tables.keys.first];
  if (sheet != null) {
    //  make horizontal room for the columns

    //  titles
    var row = sheet.maxRows + 2;
    {
      var i = 0;
      for (var title in CommitteeReportEntry.titles) {
        var cell = sheet.cell(CellIndex.indexByColumnRow(rowIndex: row, columnIndex: i++));
        cell.value = title;
        cell.cellStyle = titleCellStyle;
      }
    }
    row++;

    //  data
    for (var entry in entries) {
      entry.toXlsx(sheet, row++);
    }
  }

  await excel.encode().then((onValue) {
    File('${homePath()}/junk/excel_written.xlsx')
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);
  });

  return true;
}

class ElizabethBuildingCommittee {
  ElizabethBuildingCommittee() {
    _styles['Est. Cost'] = _dataCellRightStyle;
    _styles['Condo ID'] = _dataCellClipStyle;
    _styles['Asset ID'] = _dataCellClipStyle;
    _styles['Item ID'] = _dataCellClipStyle;
  }

  Future<bool> processFile(String filePathAsString) async {
    await input(filePathAsString);
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

  Future<bool> output() async {
    var excel = Excel.createExcel();

    var outputColumns = <int>[
      // 0,  // Blank Column
      // 1,  // Condo ID
      // 2,  // Asset ID
      3, // Item ID
      // 4,  // Condo Name
      // 5, // Asset Category
      // 6, // Unknown 20-170?
      // 7, // Unknown 10-110?
      8, // Asset Ref
      9, // Asset Name
      10, // ID
      11, // Task
      12, // Description
      13, // Frequency
      14, // Due
      // 15, // Cost high?
      // 16, // Cost low?
      // 17, // Comment
      // 18, // Unknown 2-100?
      // 19, // Always 100?
      20, // Est. Cost
    ];

    {
      excel.rename(await excel.getDefaultSheet(), 'Elizabeth');
      var sheet = excel.sheets.values.first;

      {
        var i = 0;
        for (var outputColumn in outputColumns) {
          sheet.updateCell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: i), _columnNames[outputColumn],
              cellStyle: _titleCellStyle);
          i++;
        }
      }

      var orderedRows = SplayTreeSet<_OrderedRow>();

      //  sort the rows
      for (var map in _rows) {
        orderedRows.add(_OrderedRow(map));
      }

      {
        var r = 0;
        for (var row in orderedRows) {
          var i = 0;
          for (var outputColumn in outputColumns) {
            var name = _columnNames[outputColumn];
            sheet.updateCell(CellIndex.indexByColumnRow(rowIndex: r + 1, columnIndex: i++), row.map[name],
                cellStyle: _styles[name] ?? _dataCellLeftStyle);
          }
          r++;
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
}

class _OrderedRow implements Comparable<_OrderedRow> {
  _OrderedRow(this.map)
      : _dueIndex = _dueColumnSortIndex(map['Due']),
        _frequencyIndex = _frequencyColumnSortIndex(map['Frequency']);

  final int _dueIndex;
  final int _frequencyIndex;

  var map = <String, String>{};

  @override
  int compareTo(_OrderedRow other) {
    int? ret = _dueIndex.compareTo(other._dueIndex);
    if (ret != 0) {
      return ret;
    }
    ret = _frequencyIndex.compareTo(other._frequencyIndex);
    if (ret != 0) {
      return ret;
    }

    for (var id in ['Asset Ref', 'Asset Name', 'Task', 'Description', 'Item ID']) {
      ret = map[id]?.compareTo(other.map[id] ?? '');
      if (ret != null && ret != 0) {
        return ret;
      }
    }

    return 0;
  }
}

final RegExp _numberRegexp = RegExp(r'^\d+$');

int _dueColumnSortIndex(String? s) {
  if (s == null) {
    return -1;
  }
  if (s == 'Annually') {
    return 0;
  }

  var m = _numberRegexp.firstMatch(s);
  if (m == null) {
    return -1;
  }

  return int.parse(m.group(0)!);
}

final RegExp _yrsRegexp = RegExp(r'^(\d+) Yrs$');

final _frequencyValueMap = <String, int>{
  'Daily': 1,
  'As Required': 2,
  'Weekly': 7,
  'Monthly': 30,
  'Quarterly': 365 ~/ 4,
  'Semi-Annually': 365 ~/ 2,
  'Annually': 365,
};

int _frequencyColumnSortIndex(String? s) {
  if (s == null) {
    return -1;
  }

  var ret = _frequencyValueMap[s];
  if (ret != null) {
    return ret;
  }

  var m = _yrsRegexp.firstMatch(s);
  if (m == null) {
    return -1;
  }

  return int.parse(m.group(1)!) * 365; //   units of days
}
