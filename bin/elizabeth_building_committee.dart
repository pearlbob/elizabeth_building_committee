import 'dart:io' as io;
import 'dart:math';

import 'package:elizabeth_building_committee/committe_report_entry.dart';
import 'package:gcloud/storage.dart';
import 'package:googleapis/drive/v3.dart';
import 'package:googleapis/storage/v1.dart';
import 'package:gsheets/gsheets.dart';
import 'package:csv/csv.dart';
import 'package:googleapis_auth/auth_io.dart' as auth;
import 'package:gcloud/service_scope.dart' as ss;

//  google auth credentials

var _jsonDriveCredentials = io.File(
        '${homePath()}/googleCloudService/elizabeth_building_committee_client_id.json')
    .readAsStringSync();
var _jsonGSheetCredentials = io.File(
        '${homePath()}/googleCloudService/bsteele com Project-16884817dba8.json')
    .readAsStringSync();
var _gsheets = GSheets(_jsonGSheetCredentials);
var _project = 'fit-union-164517';

// Read the service account credentials from the file.
var gSheetCredentials =
    auth.ServiceAccountCredentials.fromJson(_jsonGSheetCredentials);

var _gDriveClient;
var _gSheetClient;
Storage? _storage;

var entries = <CommitteeReportEntry>[];

void main(List<String> arguments) async {
  await _initializeClient();

  await testCSV();

  // await testGoogleStorageListBuckets();
  //await testGSheetsFolders();

  await testGSheets();

  var sb = StringBuffer(CommitteeReportEntry.csvTitles());
  sb.writeln('');
  {
    var listList = <List<String>>[];
    for (var entry in entries) {
      listList.add(entry.toList());
    }
    var convertor = ListToCsvConverter();
    sb.writeln(convertor.convert(listList));
  }

  print(sb.toString());

  io.exit(0);
}

Future<bool> _initializeClient() async {
  // Get an HTTP authenticated client using the service account credentials.
  // auth.AccessCredentials(,['https://www.googleapis.com/auth/drive']);
  // var clientCredentials = auth.AccessCredentials(_jsonDriveCredentials);
  // _gDriveClient = await auth.clientViaServiceAccount(
  //     clientCredentials,
  //     [StorageApi.DevstorageReadOnlyScope]);
  _gSheetClient = await auth
      .clientViaServiceAccount(gSheetCredentials, [...Storage.SCOPES]);
  _storage = Storage(_gSheetClient, _project);

  // await ss.fork(() async {
  //   // register the services in the new service scope.
  //   registerStorageService(_storage);
  //   // Run application using these services.
  // });
  return true;
}

enum _State {
  initial,
  vendor,
  vendorData,
}

Future<bool> testCSV() async {
  var csvAsString = io.File('./lib/assets/maintenance_test_report_20210128.csv')
      .readAsStringSync();
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

          if (entry.dateString.isEmpty &&
              entry.vendor.isEmpty &&
              entry.description.isEmpty) {
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
    print(
        '${sheet.title}  rows: ${sheet.rowCount}, cols: ${sheet.columnCount}');

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

String homePath() {
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
