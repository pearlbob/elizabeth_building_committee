import 'package:csv/csv.dart';

enum CommitteeReportEntrySource {
  FAM_report_csv,
  GSheets,
}

String CommitteeReportEntrySourceAbbreviation(
    CommitteeReportEntrySource source) {
  switch (source) {
    case CommitteeReportEntrySource.FAM_report_csv:
      return 'FAM';
    case CommitteeReportEntrySource.GSheets:
      return 'GS';
    default:
      return 'unknown: $source';
  }
}

/// Elizabeth Building Committee report entry (row)
class CommitteeReportEntry {
  //
  CommitteeReportEntry(
    this.source, {
    item = '',
    description = '',
    dateString = '',
    vendor= '',
    fullFocusWOReference= '',
    type= '',
    budgetedItem= '',
    reserve= '',
    cost= '',
    status= '',
    completed= '',
  })  : item = item,
        description = description,
        dateString = dateString,
        fullFocusWOReference = fullFocusWOReference,
        vendor = vendor,
        //description = description,
        type = type,
        budgetedItem = budgetedItem,
        reserve = reserve,
        cost = cost,
        status = status,
        completed = completed;

  @override
  String toString() {
    return 'CommitteeReportEntry{\n'
        '\tsource: ${CommitteeReportEntrySourceAbbreviation(source)},\n'
        '\tdateString: $dateString,\n'
        '\tfullFocusWOReference: $fullFocusWOReference,\n'
        '\titem: $item,\n'
        '\tvendor: $vendor,\n\tdescription: $description,\n'
        '\ttype: $type,\n\tbudgetedItem: $budgetedItem,\n'
        '\treserve: $reserve,\n\tcost: $cost,\n'
        '\tstatus: $status,\n\tcompleted: $completed\n}\n';
  }

  String toJSON() {
    var sb = StringBuffer();
    sb.writeln('{');
    sb.writeln(
        '\t"source": "${CommitteeReportEntrySourceAbbreviation(source)}",');
    sb.writeln('\t"dateString": "$dateString",');
    sb.writeln('\t"fullFocusWOReference": "$fullFocusWOReference",');
    sb.writeln('\t"vendor": "$vendor",');
    sb.writeln('\t"item": "$item",');
    sb.writeln('\t"description": "$description",');
    sb.writeln('\t"type": "$type",');
    sb.writeln('\t"budgetedItem": "$budgetedItem",');
    sb.writeln('\t"reserve": "$reserve",');
    sb.writeln('\t"cost": "$cost",');
    sb.writeln('\t"status": "$status",');
    sb.writeln('\t"completed": "$completed"');
    sb.writeln('}');
    return sb.toString();
  }

  static String csvTitles() {
    return         'Source,'
        'Date String,'
        'Full Focus WO Reference,'
        'Vendor,'
        'Item,'
        'Description,'
        'Type,'
        'Budgeted Item,'
        'Reserve,'
        'Cost,'
        'Status,'
        'Completed';
  }

  List<String> toList() {
    return [
      CommitteeReportEntrySourceAbbreviation(source),
      dateString,
      fullFocusWOReference,
      vendor,
      item,
      description,
      type,
      budgetedItem,
      reserve,
      cost,
      status,
      completed
    ];
  }

  String toCSV() {
    var sb = StringBuffer();
    sb.write('${CommitteeReportEntrySourceAbbreviation(source)},');
    sb.write('$dateString,');
    sb.write('$fullFocusWOReference,');
    sb.write('$vendor,');
    sb.write('$item,');
    sb.write('$description,');
    sb.write('$type,');
    sb.write('$budgetedItem,');
    sb.write('$reserve,');
    sb.write('$cost,');
    sb.write('$status,');
    sb.write('$completed');
    return sb.toString();
  }

  CommitteeReportEntrySource source;

  String dateString;

  String fullFocusWOReference = '';

  String vendor = '';

  String item = '';

  String description = '';

  String type = '';

  String budgetedItem = '';

  String reserve = '';

  String cost = '';

  String status = '';

  String completed = '';
}
