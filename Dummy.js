// dummy.js — manual test runner for Auto Migration
// Usage: run testAutoMigrateJlid() from GAS editor

function testAutoMigrateJlid() {
  var JLID = 'JL39611449152C2';
  Logger.log('========== AUTO MIGRATE TEST ==========');
  Logger.log('JLID: ' + JLID);
  Logger.log('Time: ' + new Date().toISOString());

  try {
    var result = runAutoMigrateForJlid(JLID, '');
    Logger.log('========== RESULT ==========');
    Logger.log(JSON.stringify(result, null, 2));
  } catch (e) {
    Logger.log('========== ERROR ==========');
    Logger.log(e.message);
    Logger.log(e.stack);
  }

  Logger.log('========== DONE ==========');
}
