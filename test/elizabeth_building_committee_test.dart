
import 'package:test/test.dart';

void main() {
  test('calculate', () {
    expect('test here', 'test here');
  });

  test('regex', () {

    var m = _sampleRegexp.firstMatch(' asdf  asdf  k12348asdfasdf dfgsdfg');

    if ( m != null ){
      print( 'm says: ${m.group(0)}, digits: ${m.group(1)}, post: ${m.group(2)}');
      var i = int.parse(m.group(1)!);
      print( 'i: $i');
    } else {
      print( 'm says: no match');
    }
  });
}

 final RegExp _sampleRegexp = RegExp(r'k(\d+)([\w]+)');