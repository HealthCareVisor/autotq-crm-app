// AutoTQ Outlook button: "Send personal copies". No window: it hands the draft to the mail-merge function, which mails the rep
// a preview of the first copy; once the rep confirms in TQbot (the CRM side panel), it sends one personal copy per person from
// the rep's own mailbox in the background. The draft is discarded after it's queued.
/* global Office */
(function () {
  var LOCAL = /^(localhost|127\.0\.0\.1)$/.test(location.hostname);
  var SUPA_URL = (LOCAL && window.MERGE_API) || "https://nrbjtqczntrwqoheurpt.supabase.co";
  var SUPA_KEY = "sb_publishable_nJhBA0zj3h11qzDWeIU0VA_1LrjKOp-";
  var SKEY = "autotq_merge_session_v1";
  var BASE = location.href.replace(/[^/]*$/, "");

  function getS() { try { return JSON.parse(localStorage.getItem(SKEY) || "null"); } catch (e) { return null; } }
  function setS(s) { try { if (s) localStorage.setItem(SKEY, JSON.stringify(s)); else localStorage.removeItem(SKEY); } catch (e) { /* */ } }
  function pget(obj, arg) {
    return new Promise(function (res, rej) {
      var cb = function (r) { r.status === "succeeded" ? res(r.value) : rej(new Error((r.error && r.error.message) || "Outlook didn't answer")); };
      arg === undefined ? obj.getAsync(cb) : obj.getAsync(arg, cb);
    });
  }
  function note(item, msg, isErr) {
    var m = String(msg).slice(0, 150);
    var o = isErr ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: m }
      : { type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage, message: m, icon: "Icon16", persistent: false };
    try { item.notificationMessages.replaceAsync("autotq", o); } catch (e) { /* */ }
  }
  function recipients(item) {
    return Promise.all([pget(item.to), pget(item.cc), pget(item.bcc)]).then(function (lists) {
      var out = [];
      lists.forEach(function (l) { (l || []).forEach(function (r) { out.push({ email: r.emailAddress || "", name: r.displayName || "", type: r.recipientType || "" }); }); });
      return out;
    });
  }
  function refresh(s) {
    if (!s || !s.refresh_token) return Promise.resolve(null);
    return fetch(SUPA_URL + "/auth/v1/token?grant_type=refresh_token", { method: "POST", headers: { apikey: SUPA_KEY, "Content-Type": "application/json" },
      body: JSON.stringify({ refresh_token: s.refresh_token }) })
      .then(function (r) { return r.ok ? r.json() : null; }).then(function (j) { setS(j); return j; }).catch(function () { return null; });
  }
  // one-time sign-in box (the CRM's own sign-in); the session comes back through messageParent
  function signIn() {
    var email = "";
    try { email = Office.context.mailbox.userProfile.emailAddress || ""; } catch (e) { /* */ }
    return new Promise(function (res, rej) {
      Office.context.ui.displayDialogAsync(BASE + "signin.html?email=" + encodeURIComponent(email) + (LOCAL && window.MERGE_API ? "&api=" + encodeURIComponent(window.MERGE_API) : ""),
        { height: 52, width: 28, displayInIframe: true }, function (r) {
          if (r.status !== "succeeded") return rej(new Error("Couldn't open the sign-in box."));
          var d = r.value;
          d.addEventHandler(Office.EventType.DialogMessageReceived, function (m) {
            d.close();
            var s = null; try { s = JSON.parse(m.message); } catch (e) { /* */ }
            if (!s || !s.access_token) return rej(new Error("Not signed in."));
            setS(s); res(s);
          });
          d.addEventHandler(Office.EventType.DialogEventReceived, function () { rej(new Error("Not signed in.")); });
        });
    });
  }
  function queue(s, payload) {
    return fetch(SUPA_URL + "/functions/v1/mail-merge", { method: "POST",
      headers: { apikey: SUPA_KEY, Authorization: "Bearer " + s.access_token, "Content-Type": "application/json" },
      body: JSON.stringify(Object.assign({ action: "queue" }, payload)) })
      .then(function (r) { return r.json().catch(function () { return {}; }).then(function (j) { return { status: r.status, j: j }; }); });
  }
  function closeDraft(item, done) {
    var can = false;
    try { can = Office.context.requirements.isSetSupported("Mailbox", "1.14") && typeof item.closeAsync === "function"; } catch (e) { /* */ }
    if (!can) return done(false);
    item.closeAsync({ discardItem: true }, function () { done(true); });
  }

  function sendCopies(event) {
    var item = Office.context.mailbox.item, finished = false;
    var finish = function () { if (!finished) { finished = true; try { event.completed(); } catch (e) { /* */ } } };
    var payload;
    note(item, "AutoTQ: getting your email ready…");
    Promise.all([pget(item.subject), pget(item.body, Office.CoercionType.Html), recipients(item)]).then(function (v) {
      var mailbox = ""; try { mailbox = Office.context.mailbox.userProfile.emailAddress || ""; } catch (e) { /* */ }
      payload = { subject: v[0] || "", html: v[1] || "", recipients: v[2], mailbox: mailbox };
      if (!payload.recipients.length) throw new Error("Add the people to To or Bcc first.");
      var s = getS();
      return s ? s : signIn();
    }).then(function (s) {
      return queue(s, payload).then(function (r) {
        if (r.status !== 401) return r;
        return refresh(s).then(function (s2) { return s2 || signIn(); }).then(function (s3) { return queue(s3, payload); });
      });
    }).then(function (r) {
      var j = r.j || {};
      if (j.wrongUser) setS(null); // signed in to the CRM as someone else: the next click asks again
      if (!j.ok) throw new Error(j.error || "Couldn't queue it (" + r.status + ").");
      var held = (j.held || []).length;
      var msg = "Preview sent to your inbox. Confirm in TQbot to send " + j.ready + " personal " + (j.ready === 1 ? "copy" : "copies") + (held ? " (" + held + " held)" : "") + ".";
      note(item, msg);
      // leave the message up a moment, then discard the draft so it can't be sent by hand with {{name}} in it
      setTimeout(function () { closeDraft(item, function (closed) { if (!closed) note(item, msg + " You can delete this draft."); finish(); }); }, 3500);
    }).catch(function (e) {
      note(item, "AutoTQ: " + ((e && e.message) || e), true);
      finish();
    });
  }

  window.sendCopies = sendCopies;
  Office.onReady(function () {
    try { if (Office.actions && Office.actions.associate) Office.actions.associate("sendCopies", sendCopies); } catch (e) { /* */ }
  });
})();
