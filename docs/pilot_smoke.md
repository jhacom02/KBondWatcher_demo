# Pilot smoke checklist (manual)
#
# Build / mode
# [ ] Source: KBOND_DEPLOY_MODE unset → deploy_mode=dev
# [ ] Frozen/Nuitka: deploy_mode=pilot even if KBOND_DEPLOY_MODE=dev
# [ ] Nuitka onefolder: VERSION.txt + demo_expiry.txt (build+7d); no .py; no admin.db in dist
# [ ] start.bat present; replace KBOND_ADMIN_URL + KBOND_SIGNING_PUBLIC_KEY before zip
# [ ] Frozen START spawns main.exe --run-profile (not missing main.py)
#
# Credential / lease / profile
# [ ] device.json has credential_blob + credential_protection=dpapi (no plaintext credential)
# [ ] Dev: local issue_lease works; Pilot: local issue_lease raises
# [ ] Lease TTL default 7d; Admin refuses reissue after pilot_expires_at (absolute window)
# [ ] Lease binds device_id+trader_id+profile_version+min_engine+expires_at+enabled (no exact engine_version)
# [ ] Signature covers policy_payload only; mutable change keeps sig valid; locked tamper fails START
# [ ] Locked fields (name/chat/mode/loop/excel workbook·sheet/template): draft Submit→Approve; runtime fields (instrument/looking/qty/threshold/yield_prefix/input·output cell): STOP→Save without re-approve
# [ ] After engine upgrade from full-profile sig: re-Submit/Approve once (no dual-verify)
# [ ] PROFILE_RUNTIME_SAVED / START·STOP·ERROR audit include prefs snapshot
# [ ] Tampered locked fields or .sig fails START + PROFILE_REJECTED / LICENSE_REJECTED audit
# [ ] Past demo_expiry.txt date blocks START / policy poll soft STOP
# [ ] No KBOND_SIGNING_PUBLIC_KEY (and no public key file) → verify fails immediately
#
# Controller policy poll (not in Quote→Excel→Send)
# [ ] START with valid signed profile + lease
# [ ] Admin disable device/trader while WATCHING → soft STOP within policy_poll_seconds (~60s)
# [ ] Lease expiry / min_engine bump → soft STOP
# [ ] Admin DOWN: continues on cached lease until expiry; then STOP / block START
#
# Admin remote
# [ ] KBOND_ADMIN_URL HTTPS in pilot; device register + HMAC auth
# [ ] UI shows last_seen / last_lease / last_audit / pilot_expires / credential_protection / disable
# [ ] Audit ingest UNIQUE event_id; duplicates counted; client retries + stale status visible
#
# Runtime regression
# [ ] python main.py --serve ; calibrate; START/STOP
# [ ] MODE 1 / 2 / 3 quote path; Excel close → ERROR (no reconnect)
# [ ] Tab close / refresh → STOP via keepalive
# [ ] Ambiguous quotes / sanity band → ERROR, no send
# [ ] NO_TRIGGER skips send and stays WATCHING
# [ ] sent_after=loop (demo) re-watches after successful send; mode 1 forces exit
# [ ] Stale watcher PID cleared on START
# [ ] Closing serve console / Ctrl+C stops watcher (job + shutdown hook)
# [ ] pytest -q green
