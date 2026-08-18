# Pilot smoke checklist (manual)
#
# Build / mode
# [ ] Source: KBOND_DEPLOY_MODE unset → deploy_mode=dev
# [ ] Frozen/Nuitka: deploy_mode=pilot even if KBOND_DEPLOY_MODE=dev
# [ ] Nuitka onefolder: VERSION.txt + demo_expiry.txt present; no .py left in dist
#
# Credential / lease / profile
# [ ] device.json has credential_blob + credential_protection=dpapi (no plaintext credential)
# [ ] Dev: local issue_lease works; Pilot: local issue_lease raises
# [ ] Lease binds device_id+trader_id+profile_version+min_engine+expires_at+enabled (no exact engine_version)
# [ ] Profile: Local Web saves draft only when Admin URL set; submit → Admin approve → apply
# [ ] Tampered profile.json or .sig fails START + PROFILE_REJECTED / LICENSE_REJECTED audit
#
# Controller policy poll (not in Quote→Excel→Send)
# [ ] START with valid signed profile + lease
# [ ] Admin disable device/trader while WATCHING → soft STOP within policy_poll_seconds (~60s)
# [ ] Lease expiry / min_engine bump → soft STOP
# [ ] Admin DOWN: continues on cached lease until expiry; then STOP / block START
#
# Admin remote
# [ ] KBOND_ADMIN_URL HTTPS in pilot; device register + HMAC auth
# [ ] UI shows last_seen / last_lease / last_audit / credential_protection / disable
# [ ] Audit ingest UNIQUE event_id; duplicates counted; client retries + stale status visible
#
# Runtime regression
# [ ] python main.py --serve ; calibrate; START/STOP
# [ ] MODE 1 / 2 / 3 quote path; Excel EXCEL_WAIT reconnect
# [ ] Ambiguous quotes fail-closed; send focus/clipboard guards
# [ ] Stale watcher PID cleared on START
# [ ] pytest -q green
