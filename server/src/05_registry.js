// ========== API action registry ==========
// Every handler module declares its own actions with defineAction(); doGet/doPost
// dispatch through the registry instead of a 300-line switch, so adding an
// action is one line next to its handler and the auth rule lives with it.
//   defineAction('startExam', { methods: ['POST'], auth: 'examinee', handler: handleStartExam });
//   auth: 'none' | 'examiner' | 'teacher' | 'examinee' | 'gateway'
//   methods: any of 'GET', 'POST' (a GET to a POST-only action is refused).
//   rateLimit: optional { max, windowSec, id: function(p) -> identifier }
// The registry is kept on the function object so module load order does not
// matter (a module may register before this file's vars would have run).
function apiRegistry() {
  if (!apiRegistry._actions) apiRegistry._actions = {};
  return apiRegistry._actions;
}
function defineAction(name, spec) {
  if (!name || !spec || typeof spec.handler !== 'function') throw new Error('defineAction: bad spec for ' + name);
  apiRegistry()[name] = {
    name: name,
    methods: spec.methods || ['GET'],
    auth: spec.auth || 'none',
    handler: spec.handler,
    rateLimit: spec.rateLimit || null
  };
}
function apiActionNames() { return Object.keys(apiRegistry()).sort(); }
