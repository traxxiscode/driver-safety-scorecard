/**
 * Geotab HOS Alert Emailer Add-in
 */
geotab.addin.hosAlerter = function () {
    'use strict';

    var api, state, elAddin;
    var docRef = null; // Firestore document reference for this database

    // ── Firebase init ───────────────────────────────────────────────────────────

    var firebaseConfig = {
        apiKey:            "AIzaSyCDquA4ZS0rGVpwwMp-e9g0hK4Rnp8Aqxs",
        authDomain:        "hos-volations.firebaseapp.com",
        projectId:         "hos-volations",
        storageBucket:     "hos-volations.firebasestorage.app",
        messagingSenderId: "775730536667",
        appId:             "1:775730536667:web:008b434fc859bb3e232cfe"
    };

    if (!firebase.apps.length) {
        firebase.initializeApp(firebaseConfig);
    }
    var db = firebase.firestore();

    // ── Initialise: get session → upsert Firestore doc → load recipients ────────

    function init() {
        api.getSession(function (session) {
            var databaseName = session.database;

            document.getElementById('currentDatabase').textContent = databaseName;

            firebase.auth().signInAnonymously().then(function () {

                // Use database name as the document ID — simple and deterministic
                docRef = db.collection('hos_configurations').doc(databaseName);

                docRef.get().then(function (snap) {
                    if (!snap.exists) {
                        return docRef.set({
                            database_name: databaseName,
                            recipients:    [],
                            active:        true,
                            created_at:    firebase.firestore.FieldValue.serverTimestamp()
                        });
                    }
                }).then(function () {
                    loadRecipients();
                }).catch(function (err) {
                    showAlert('Firestore error: ' + err.message, 'danger');
                    hideLoading();
                });

            }).catch(function (err) {
                showAlert('Auth error: ' + err.message, 'danger');
                hideLoading();
            });
        });
    }

    // ── Firestore: load ─────────────────────────────────────────────────────────

    function loadRecipients() {
        docRef.get().then(function (snap) {
            var recipients = snap.exists ? (snap.data().recipients || []) : [];
            renderRecipients(recipients);
            document.getElementById('recipientCount').textContent = recipients.length;
        }).catch(function (err) {
            showAlert('Error loading recipients: ' + err.message, 'danger');
        }).finally(function () {
            hideLoading();
        });
    }

    // ── Firestore: add ──────────────────────────────────────────────────────────

    function addRecipient(email) {
        setButtonLoading('addRecipientBtn', true);

        docRef.get().then(function (snap) {
            var recipients = snap.exists ? (snap.data().recipients || []) : [];

            if (recipients.some(function (r) { return r.email === email; })) {
                showAlert('That email is already in the list.', 'warning');
                return;
            }

            recipients.push({ email: email, added_at: new Date().toISOString() });

            return docRef.update({
                recipients: recipients,
                updated_at: firebase.firestore.FieldValue.serverTimestamp()
            }).then(function () {
                showAlert(email + ' added.', 'success');
                document.getElementById('recipientEmail').value = '';
                loadRecipients();
            });

        }).catch(function (err) {
            showAlert('Error adding recipient: ' + err.message, 'danger');
        }).finally(function () {
            setButtonLoading('addRecipientBtn', false);
        });
    }

    // ── Firestore: remove ───────────────────────────────────────────────────────

    function removeRecipient(email) {
        docRef.get().then(function (snap) {
            var recipients = snap.exists ? (snap.data().recipients || []) : [];
            var updated    = recipients.filter(function (r) { return r.email !== email; });

            return docRef.update({
                recipients: updated,
                updated_at: firebase.firestore.FieldValue.serverTimestamp()
            }).then(function () {
                showAlert(email + ' removed.', 'success');
                loadRecipients();
            });

        }).catch(function (err) {
            showAlert('Error removing recipient: ' + err.message, 'danger');
        });
    }

    // ── Render ──────────────────────────────────────────────────────────────────

    function renderRecipients(recipients) {
        var container = document.getElementById('recipientsList');

        if (recipients.length === 0) {
            container.innerHTML =
                '<div class="empty-state">' +
                  '<i class="fas fa-inbox"></i>' +
                  '<p class="mb-1 fw-bold">No recipients yet</p>' +
                  '<small>Add an email address above to start receiving HOS alerts.</small>' +
                '</div>';
            return;
        }

        container.innerHTML = recipients.map(function (r) {
            return '<div class="recipient-item">' +
                     '<div class="recipient-email">' + r.email + '</div>' +
                     '<button class="btn btn-outline-danger btn-sm" ' +
                       'onclick="hosAlerterRemove(\'' + r.email + '\')">' +
                       '<i class="fas fa-trash me-1"></i>Remove' +
                     '</button>' +
                   '</div>';
        }).join('');
    }

    // ── UI helpers ──────────────────────────────────────────────────────────────

    function showAlert(message, type) {
        var container = document.getElementById('alertContainer');
        var id        = 'alert-' + Date.now();
        var icons     = { success: 'check-circle', danger: 'exclamation-triangle',
                          warning: 'exclamation-triangle', info: 'info-circle' };

        container.insertAdjacentHTML('beforeend',
            '<div class="alert alert-' + type + ' alert-dismissible fade show" id="' + id + '">' +
              '<i class="fas fa-' + (icons[type] || 'info-circle') + ' me-2"></i>' +
              message +
              '<button type="button" class="btn-close" data-bs-dismiss="alert"></button>' +
            '</div>'
        );

        setTimeout(function () {
            var el = document.getElementById(id);
            if (el) el.remove();
        }, 5000);
    }

    function setButtonLoading(id, loading) {
        var btn  = document.getElementById(id);
        if (!btn) return;
        var text = btn.querySelector('.btn-text');
        var spin = btn.querySelector('.btn-loading-text');
        btn.disabled = loading;
        if (text) text.style.display = loading ? 'none'        : '';
        if (spin) spin.style.display = loading ? 'inline-flex' : 'none';
    }

    function hideLoading() {
        var el = document.getElementById('initialLoadingOverlay');
        if (el) el.style.display = 'none';
    }

    // ── Globals (called from HTML onclick) ──────────────────────────────────────

    window.hosAlerterRemove = function (email) {
        if (confirm('Remove ' + email + ' from HOS alerts?')) {
            removeRecipient(email);
        }
    };

    window.hosAlerterRefresh = function () {
        if (docRef) loadRecipients();
    };

    // ── Geotab lifecycle ────────────────────────────────────────────────────────

    return {
        initialize: function (freshApi, freshState, cb) {
            api      = freshApi;
            state    = freshState;
            elAddin  = document.getElementById('hosAlerter');
            if (state.translate) state.translate(elAddin || '');
            cb();
        },

        focus: function (freshApi, freshState) {
            api   = freshApi;
            state = freshState;

            // Wire up the form submit
            var form = document.getElementById('addRecipientForm');
            if (form && !form._bound) {
                form._bound = true;
                form.addEventListener('submit', function (e) {
                    e.preventDefault();
                    var email = document.getElementById('recipientEmail').value.trim();
                    if (email) addRecipient(email);
                });
            }

            if (elAddin) elAddin.style.display = 'block';
            init();
        },

        blur: function () {
            if (elAddin) elAddin.style.display = 'none';
        }
    };
};