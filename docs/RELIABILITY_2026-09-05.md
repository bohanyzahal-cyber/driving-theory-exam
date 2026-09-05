# Reliability release — 2026-09-05-r1

This release addresses cache churn, repeated spreadsheet reads and browser requests that could remain pending after response headers arrived. The changes preserve the existing question selection, server scoring and exam registration protocol.

## Changes

- Question banks and license pools use gzip/base64 records with generation checks. Each value stays below 81 KB. Seven banks, 35 pools and the packed translation index reserve at most 423 cache keys, instead of approximately 1,700 individual translation keys alone.
- Warmup reads each language bank once per invocation, reuses the data, renews the bank TTL, removes legacy entries and verifies the complete cache after writing all pools. Cold builders claim an owned lease under a short script lock; competing requests receive a retryable response instead of sleeping inside Apps Script. The script lock is released before reading Drive or processing questions.
- Question request limits apply to each authenticated examinee within a session. Guest limits remain unchanged.
- Result submission reuses spreadsheet snapshots with targeted freshness checks and retains the late full result read for duplicate detection. It avoids loading question text for a perfect score. Full history remains available for old results and retakes.
- Client deadlines cover response bodies as well as headers. Dashboard, approval and DQ polling recover from errors without overlapping their own requests. A question-loading failure offers a cooldown-aware retry that preserves registration. Result retries retain unconfirmed data, cancel obsolete timers and distinguish different attempts on a shared device.
- `action=health&origin=examinee-app` returns the API build marker without reading Sheets or Drive. API timing logs contain the build, method, an allowlisted action name and elapsed milliseconds; request parameters are not included.

## Offline validation

Run from the repository root with Node.js:

```text
node --check external_exam_apps_script.js
node --test tests/api_reliability.test.cjs tests/cache_reliability.test.cjs tests/client_reliability.test.cjs tests/database_reliability.test.cjs
```

The database suite compares synthetic outcomes and complete result rows with commit `232ed3c`. Its large fixture (4,000 pending rows, 4,000 exam rows and 5,000 result rows) reduces full spreadsheet reads from 9 to 4 and requested cells from 769,184 to 419,142. These are workload counts, not production latency measurements.

The cache suite covers capacity, UTF-8 value sizes, partial eviction, mixed generations, parallel builder ownership, failures, TTL renewal and migration from a saturated legacy cache. Optional validation against local private generated banks prints counts only:

```text
node tests/cache_reliability.test.cjs deployment/generated
```

With the available seven local banks, verification found 1,700 unique IDs, 338 persistent cache keys, a maximum value of 80,037 bytes and no missing or mixed-generation keys. Private question data is not included in the tests or this release.

Browser smoke testing uses a separate loopback API with synthetic questions. It checks registration, approval, a simulated cache-busy response, an explicit retry, 30 answers and confirmed submission. This does not validate the production Apps Script deployment or actual classroom concurrency.

## Deploy and verify

GitHub Pages and Apps Script deploy separately. The client update remains compatible with the previous API response format.

1. Update the existing Apps Script source with the complete `external_exam_apps_script.js`. Keep the existing private answer-key file and script properties.
2. Publish a new version of the existing web-app deployment so its `/exec` URL remains unchanged.
3. Check that the public health response reports `build: "2026-09-05-r1"`. A successful Pages deployment does not prove this server step occurred.
4. Run `warmupQuestionCaches` in the Apps Script editor. Its summary must contain no `ERROR` or `cached=false`. Final cache verification must show `ready: true` and `missingOrMixedKeys: 0`. `questionCacheStatus` is also available as a read-only editor function. CacheService may evict entries later; this check is a snapshot.
5. Confirm the existing warmup trigger runs every four hours. Before the exam day, perform a complete controlled exam on two devices and verify that the examiner receives the result. Check API execution logs during that test, including `getExamQuestions`, `registerExamQuestions` and `submitResult`.

Do not use a large production load test as a substitute for a controlled end-to-end test. Local checks cannot establish production quotas, current deployment configuration, network latency or the cause of each historical timeout.

## Recovery

The previous API is retained in Apps Script deployment history and the previous client code is in Git. To roll back, restore the prior deployment and revert this release's client changes. The `qv2_` cache namespace is separate from the previous format; warm the selected version after changing deployments. Do not delete stored pending results or regenerate answer keys as part of a reliability rollback.
