// ========== Time triggers — the list, and the way to get rid of them =========
//
// Apps Script keeps running a trigger even after the function it names is gone
// from the script, and every run burns from the 90-minutes-a-day trigger budget
// of the ACCOUNT (the warmup incident: 300 s of a 361 s kill). So the names of
// the retired jobs are kept here for as long as an installation may still carry
// their triggers, and both halves of the split (DESIGN §13.3) can see them:
// this module is `both`, installNightlyJobs stays with the archive in the
// reports deployment, uninstallNightlyJobs is run in the exam deployment.
var NIGHTLY_OBSOLETE_HANDLERS = ['archiveOldPendingRows', 'archiveSheets', 'warmupQuestionCaches',
  'ensureQuestionCachesWarm', 'rebuildMissingQuestionCaches', 'rebuildAtRiskCache'];

// Run ONCE from the editor of the EXAM project after the split is deployed
// (DESIGN §13.3 step ג). The nightly jobs — archiveSheets 01:00 and
// rebuildAtRiskCache 03:00 — move to the reports project, where their handlers
// now live; a trigger left behind here would fire a function this deployment no
// longer contains, fail every night, and still be charged for.
// Safe to run in any deployment and safe to run twice: it only deletes.
function uninstallNightlyJobs() {
  var wanted = NIGHTLY_OBSOLETE_HANDLERS.slice();
  // Named again explicitly: these two are the LIVE jobs, and the guarantee of
  // this function must not depend on them happening to be in the obsolete list.
  var live = ['archiveSheets', 'rebuildAtRiskCache'];
  for (var l = 0; l < live.length; l++) if (wanted.indexOf(live[l]) === -1) wanted.push(live[l]);
  var trigs = ScriptApp.getProjectTriggers(), removed = [];
  for (var i = 0; i < trigs.length; i++) {
    var fn = trigs[i].getHandlerFunction();
    if (wanted.indexOf(fn) === -1) continue;
    ScriptApp.deleteTrigger(trigs[i]);
    removed.push(fn);
  }
  var msg = 'uninstallNightlyJobs: removed ' + removed.length + ' trigger(s) [' + removed.join(', ') +
    ']; ' + (trigs.length - removed.length) + ' other trigger(s) left untouched';
  Logger.log(msg);
  return msg;
}
