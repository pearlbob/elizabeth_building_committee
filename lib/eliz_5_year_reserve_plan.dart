import 'dart:io';

import 'package:elizabeth_building_committee/util/app_logger.dart';

class Elizabeth5YearReservePlan {
  Elizabeth5YearReservePlan(File file) {
    logger.i('Eliz5YearReservePlan(${file.path})');
    if ( !file.existsSync()){
      logger.e('file missing: ${file.path}');
      return;
    }
    logger.e('file here: ${file.absolute.path}');
  }
}
