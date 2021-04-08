import 'dart:io';

String homePath() {
  var home = '';
  var envVars = Platform.environment;
  if (Platform.isMacOS) {
    home = envVars['HOME'] ?? '';
  } else if (Platform.isLinux) {
    home = envVars['HOME'] ?? '';
  } else if (Platform.isWindows) {
    home = envVars['UserProfile'] ?? '';
  }
  return home;
}