import 'dart:io' as io;
import 'dart:math';

import 'package:elizabeth_building_committee/committe_report_entry.dart';
import 'package:excel/excel.dart';
import 'package:gcloud/storage.dart';
import 'package:googleapis/drive/v3.dart';
import 'package:googleapis_auth/auth_io.dart' as auth;
import 'package:gsheets/gsheets.dart';

//  google auth credentials
var _jsonGSheetCredentials =
io.File('${_homePath()}/googleCloudService/bsteele com Project-16884817dba8.json').readAsStringSync();
var _gsheets = GSheets(_jsonGSheetCredentials);
var _project = 'fit-union-164517';

// Read the service account credentials from the file.
var gSheetCredentials = auth.ServiceAccountCredentials.fromJson(_jsonGSheetCredentials);

var _gDriveClient;
var _gSheetClient;
Storage? _storage;

var entries = <CommitteeReportEntry>[];

void main(List<String> arguments) async {
  await _initializeClient();

  // await testGoogleStorageListBuckets();
  //await testGSheetsFolders();

  //await testGSheets();
  //await testXlsx();

  // var sb = StringBuffer(CommitteeReportEntry.csvTitles());
  // sb.writeln('');
  // {
  //   var listList = <List<String>>[];
  //   for (var entry in entries) {
  //     listList.add(entry.toList());
  //   }
  //   var convertor = ListToCsvConverter();
  //   sb.writeln(convertor.convert(listList));
  // }
  //
  // print(sb.toString());

  io.exit(0);
}

Future<bool> _initializeClient() async {
  // Get an HTTP authenticated client using the service account credentials.
  // auth.AccessCredentials(,['https://www.googleapis.com/auth/drive']);
  // var clientCredentials = auth.AccessCredentials(_jsonDriveCredentials);
  // _gDriveClient = await auth.clientViaServiceAccount(
  //     clientCredentials,
  //     [StorageApi.DevstorageReadOnlyScope]);
  _gSheetClient = await auth.clientViaServiceAccount(gSheetCredentials, [...Storage.SCOPES]);
  _storage = Storage(_gSheetClient, _project);

  // await ss.fork(() async {
  //   // register the services in the new service scope.
  //   registerStorageService(_storage);
  //   // Run application using these services.
  // });
  return true;
}

Future<bool> testGoogleStorageListBuckets() async {
  await for (var name in _storage!.listBucketNames()) {
    print('name: $name');
  }

  // var e = await _storage.bucketExists('fit-union-164517_cloudbuild');
  // print('exists: $e');

  return true;
}

Future<bool> testGSheetsFolders() async {
  // var list = files.runtimeType;
  var drive = DriveApi(_gDriveClient);
  var fileList = await drive.files.list();
  for (var f in fileList.files) {
    print('file: "${f.name}", id: ${f.id}');
  }

  return true;
}

Future<bool> testGSheets() async {
  //  find local file reference to gcloud document
  // var fileName = 'Elizabeth_building_committee_202102.gdsheet';
  // var file = File('${homePath()}/junk/$fileName');
  // var exists = file.existsSync();
  // print('elizabeth_building_committee: $file exists: $exists');

  var testId = r'1IR3i4iSRNZISY2Fnwgyg4QZCLCHdtsF7G_jwdyCNeSU';
  // https://docs.google.com/spreadsheets/d/1IR3i4iSRNZISY2Fnwgyg4QZCLCHdtsF7G_jwdyCNeSU/edit?usp=sharing
  var spreadsheet = await _gsheets.spreadsheet(testId);

  for (var sheet in spreadsheet.sheets) {
    print('');
    print('${sheet.title}  rows: ${sheet.rowCount}, cols: ${sheet.columnCount}');

    {
      var cells = sheet.cells;
      var map = cells.map;
      print(map.runtimeType);
    }

    // var allRows = await cells.allRows();
    // var map = cells.map;
    // var row  = await map.row(1);
    // print( row );

    var row = await sheet.values.row(1);
    print(row);
    for (var r = 1; r <= min(20, sheet.rowCount); r++) {
      var row = await sheet.values.row(r);
      var entry = CommitteeReportEntry(
        CommitteeReportEntrySource.GSheets,
        item: row[0],
        vendor: row[1],
        description: row[2],
        cost: row[3],
        dateString: row[4],
        completed: row[5],
        status: row[6],
      );
      entries.add(entry);
    }
  }
  print('done: $spreadsheet');
  return true;
}

Future<bool> testXlsx() async {
  var fileName = '${_homePath()}/junk/excel_file.xlsx';
  var bytes = io.File(fileName).readAsBytesSync();
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
    io.File('${_homePath()}/junk/excel_written.xlsx')
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);
  });

  return true;
}

String _homePath() {
  var home = '';
  var envVars = io.Platform.environment;
  if (io.Platform.isMacOS) {
    home = envVars['HOME'] ?? '';
  } else if (io.Platform.isLinux) {
    home = envVars['HOME'] ?? '';
  } else if (io.Platform.isWindows) {
    home = envVars['UserProfile'] ?? '';
  }
  return home;
}
