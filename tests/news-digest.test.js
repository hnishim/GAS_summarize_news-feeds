'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const SOURCE_PATH = path.join(__dirname, '..', 'GAS_news-digest_standalone.js');

function formatTokyoDate(date, timeZone) {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone,
    year: 'numeric',
    month: 'numeric',
    day: 'numeric'
  }).formatToParts(date);
  const values = Object.fromEntries(parts.map(part => [part.type, part.value]));
  return `${values.year}/${values.month}/${values.day}`;
}

function loadStandalone() {
  const source = fs.readFileSync(SOURCE_PATH, 'utf8');
  const propertyValues = new Map();
  const logs = [];
  const scriptProperties = {
    getProperty(key) {
      return propertyValues.has(key) ? propertyValues.get(key) : null;
    },
    setProperty(key, value) {
      propertyValues.set(key, String(value));
      return this;
    }
  };

  const context = {
    console,
    Logger: {
      log(message) {
        logs.push(String(message));
      }
    },
    PropertiesService: {
      getScriptProperties() {
        return scriptProperties;
      }
    },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      Charset: { UTF_8: 'UTF_8' },
      computeDigest(_algorithm, value) {
        return Array.from(crypto.createHash('sha256').update(String(value), 'utf8').digest());
      },
      base64EncodeWebSafe(bytes) {
        return Buffer.from(bytes).toString('base64url');
      },
      formatDate(date, timeZone) {
        return formatTokyoDate(date, timeZone);
      },
      sleep() {}
    },
    GmailApp: { search() { return []; } },
    UrlFetchApp: {
      fetch() {
        throw new Error('Unexpected UrlFetchApp.fetch call in repository test');
      },
      fetchAll() {
        throw new Error('Unexpected UrlFetchApp.fetchAll call in repository test');
      }
    },
    LockService: {
      getScriptLock() {
        return {
          tryLock() { return true; },
          waitLock() {},
          releaseLock() {}
        };
      }
    },
    XmlService: {
      parse() {
        throw new Error('Unexpected XmlService.parse call in repository test');
      }
    }
  };

  vm.createContext(context);
  vm.runInContext(
    `${source}\n;globalThis.__testExports = {\n` +
      `NEWS_DIGEST_CONFIG, NewsDigestPipeline, StructuredGeminiService, ` +
      `RssStateStore, processRssNews, fetchAllRssFeeds, parseFeed, processUnreadNewsletters, findUnreadEmailMessages, ` +
      `selectDryRunRssItems, findNewestItem, resolveUrl, validateSummary, ` +
      `formatSlackMessage, cleanHtml, stripCodeFence, isRetryableHttpStatus\n` +
    `};`,
    context,
    { filename: SOURCE_PATH }
  );

  return { source, context, exports: context.__testExports, propertyValues, logs };
}

function geminiResponse(result) {
  return JSON.stringify({
    candidates: [{ content: { parts: [{ text: JSON.stringify(result) }] } }]
  });
}

test('standalone source has no Spreadsheet container dependency', () => {
  const { source } = loadStandalone();
  assert.doesNotMatch(source, /\bSpreadsheetApp\b/);
  assert.doesNotMatch(source, /getActiveSpreadsheet\s*\(/);
  assert.match(source, /function\s+runNewsDigest\s*\(/);
});

test('RSS state initialization treats existing items as processed and keeps bounded history', () => {
  const { exports, propertyValues } = loadStandalone();
  const { RssStateStore, NEWS_DIGEST_CONFIG } = exports;
  const state = new RssStateStore();
  const input = Array.from({ length: 60 }, (_, index) => `item-${index}`);

  state.initialize('source-a', input.concat('item-0'));

  assert.equal(state.isInitialized('source-a'), true);
  assert.equal(state.hasProcessed('source-a', 'item-0'), true);
  assert.equal(state.hasProcessed('source-a', 'item-49'), true);
  assert.equal(state.hasProcessed('source-a', 'item-50'), false);

  state.markProcessed('source-a', 'new-item');
  const stored = JSON.parse(
    propertyValues.get(`${NEWS_DIGEST_CONFIG.PROPERTIES.RSS_STATE_PREFIX}source-a`)
  );
  assert.equal(stored[0], 'new-item');
  assert.equal(stored.length, NEWS_DIGEST_CONFIG.RSS.STORED_IDS_PER_SOURCE);
});

test('pipeline posts relevant items and filters irrelevant items', () => {
  const { exports } = loadStandalone();
  const { NewsDigestPipeline } = exports;
  const sent = [];
  const item = {
    type: 'rss',
    id: '1',
    sourceId: 'source',
    source: 'Example Pharma',
    title: 'Antibody trial starts',
    content: 'content',
    url: 'https://example.com/article',
    publishedAt: new Date('2026-09-13T00:00:00+09:00')
  };

  const relevant = new NewsDigestPipeline(
    { classify() { return { relevant: true, summary: '抗体医薬の臨床試験を開始した。' }; } },
    { send(message) { sent.push(message); } }
  );
  assert.equal(relevant.process(item), 'posted');
  assert.equal(sent.length, 1);
  assert.match(sent[0], /\*Example Pharma,\* 2026\/9\/13/);
  assert.match(sent[0], /抗体医薬の臨床試験を開始した。/);
  assert.match(sent[0], /https:\/\/example\.com\/article/);

  const irrelevant = new NewsDigestPipeline(
    { classify() { return { relevant: false, summary: '' }; } },
    { send(message) { sent.push(message); } }
  );
  assert.equal(irrelevant.process(item), 'filtered');
  assert.equal(sent.length, 1);
});

test('Gemini structured result preserves relevant/irrelevant contract', () => {
  const { exports, propertyValues } = loadStandalone();
  const { StructuredGeminiService } = exports;
  propertyValues.set('GEMINI_API_KEY', 'test-key');
  const service = new StructuredGeminiService();

  service.requestWithRetry = () => geminiResponse({ relevant: true, summary: '  対象要約  ' });
  const relevant = service.classify({
    type: 'rss', source: 'source', title: 'title', url: 'https://example.com', content: 'body'
  });
  assert.equal(relevant.relevant, true);
  assert.equal(relevant.summary, '対象要約');

  service.requestWithRetry = () => geminiResponse({
    relevant: false,
    summary: 'モデルが返した不要な文字列'
  });
  const irrelevant = service.classify({
    type: 'rss', source: 'source', title: 'title', url: 'https://example.com', content: 'body'
  });
  assert.equal(irrelevant.relevant, false);
  assert.equal(irrelevant.summary, '');
});

test('summary validation enforces non-empty and 200-character maximum', () => {
  const { validateSummary } = loadStandalone().exports;
  assert.doesNotThrow(() => validateSummary('あ'.repeat(200)));
  assert.throws(() => validateSummary(''), /要約が空/);
  assert.throws(() => validateSummary('あ'.repeat(201)), /200字を超えています/);
});

test('RSS URL resolution covers absolute, protocol-relative, root and relative URLs', () => {
  const { resolveUrl } = loadStandalone().exports;
  const base = 'https://example.com/news/feed.xml';
  assert.equal(resolveUrl('https://other.example/a', base), 'https://other.example/a');
  assert.equal(resolveUrl('//cdn.example/a', base), 'https://cdn.example/a');
  assert.equal(resolveUrl('/release/1', base), 'https://example.com/release/1');
  assert.equal(resolveUrl('release/2', base), 'https://example.com/news/release/2');
});

test('email discovery selects unread messages oldest-first and respects the limit', () => {
  const { exports, context } = loadStandalone();
  const makeMessage = (id, time, unread) => ({
    getId() { return id; },
    getDate() { return new Date(time); },
    isUnread() { return unread; }
  });
  const newest = makeMessage('newest', '2026-09-13T03:00:00Z', true);
  const oldest = makeMessage('oldest', '2026-09-13T01:00:00Z', true);
  const middle = makeMessage('middle', '2026-09-13T02:00:00Z', true);
  const read = makeMessage('read', '2026-09-13T00:00:00Z', false);

  context.GmailApp.search = () => [
    { getMessages() { return [newest, read]; } },
    { getMessages() { return [middle, oldest]; } }
  ];

  const selected = exports.findUnreadEmailMessages('query', 2);
  assert.equal(selected.length, 2);
  assert.equal(selected[0].getId(), 'oldest');
  assert.equal(selected[1].getId(), 'middle');
});

test('successfully handled email is marked read', () => {
  const { exports, context } = loadStandalone();
  let markedRead = 0;
  const message = {
    getId() { return 'mail-1'; },
    getFrom() { return 'sender@example.com'; },
    getSubject() { return 'subject'; },
    getPlainBody() { return 'body'; },
    getDate() { return new Date('2026-09-13T00:00:00Z'); },
    isUnread() { return true; },
    markRead() { markedRead++; }
  };
  context.GmailApp.search = () => [{ getMessages() { return [message]; } }];

  exports.processUnreadNewsletters({ process() { return 'filtered'; } });
  assert.equal(markedRead, 1);
});

test('dry-run RSS selection takes the newest item per healthy source without mutating input', () => {
  const { selectDryRunRssItems } = loadStandalone().exports;
  const older = { id: 'older', publishedAt: new Date('2026-09-10T00:00:00Z') };
  const newer = { id: 'newer', publishedAt: new Date('2026-09-11T00:00:00Z') };
  const items = [older, newer];
  const results = [
    { source: { name: 'A' }, items, error: null },
    { source: { name: 'B' }, items: [{ id: 'b', publishedAt: new Date('2026-09-12T00:00:00Z') }], error: null },
    { source: { name: 'C' }, items: [], error: 'HTTP 500' }
  ];

  const selected = selectDryRunRssItems(results);
  assert.equal(selected[0].id, 'newer');
  assert.equal(selected[1].id, 'b');
  assert.deepEqual(items.map(item => item.id), ['older', 'newer']);
});

test('helper behavior needed by external-service adapters remains bounded', () => {
  const { cleanHtml, stripCodeFence, isRetryableHttpStatus } = loadStandalone().exports;
  assert.equal(cleanHtml('<p>Hello &amp; world</p><script>bad()</script>'), 'Hello & world');
  assert.equal(stripCodeFence('```json\n{"ok":true}\n```'), '{"ok":true}');
  assert.equal(isRetryableHttpStatus(429), true);
  assert.equal(isRetryableHttpStatus(503), true);
  assert.equal(isRetryableHttpStatus(400), false);
});

/*
 * HIR-322: 既存候補に対して先行作成する振る舞いテスト。
 * GAS の外部サービスはここでは模擬し、実環境の成功を主張しない。
 */

test('RSS history does not re-post articles older than its bounded selection window', () => {
  const { exports, context, propertyValues } = loadStandalone();
  const { RssStateStore, NEWS_DIGEST_CONFIG, processRssNews } = exports;
  const source = { id: 'test-rss', name: 'Test RSS', feedUrl: 'https://example.test/rss' };
  const now = Date.parse('2026-09-25T00:00:00Z');
  const article = (index, id) => ({
    type: 'rss', id: id || 'article-' + index, sourceId: source.id,
    source: source.name, title: 'Article ' + index, content: 'content',
    url: 'https://example.test/article/' + index,
    publishedAt: new Date(now - index * 60000)
  });
  const items = Array.from({ length: 55 }, (_, index) => article(index));
  const store = new RssStateStore();
  store.initialize(source.id, items.map(item => item.id));
  context.__feedResults = [{ source, items, error: null }];
  vm.runInContext('fetchAllRssFeeds = () => globalThis.__feedResults', context);
  const posted = [];
  const pipeline = { process(item) { posted.push(item.id); return 'posted'; } };

  processRssNews(pipeline);
  assert.deepEqual(posted, [], 'expired state IDs are not newly published articles');
  assert.equal(JSON.parse(propertyValues.get(
    NEWS_DIGEST_CONFIG.PROPERTIES.RSS_STATE_PREFIX + source.id
  )).length, NEWS_DIGEST_CONFIG.RSS.STORED_IDS_PER_SOURCE);

  const newItem = article(-1, 'genuinely-new');
  context.__feedResults = [{ source, items: [newItem, ...items], error: null }];
  processRssNews(pipeline);
  processRssNews(pipeline);
  assert.deepEqual(posted, ['genuinely-new'], 'new item is posted once; old items never replay');
});

test('RSS batch failure falls back to individual sources; one failure does not hide healthy results', () => {
  const { exports, context } = loadStandalone();
  vm.runInContext('RSS_SOURCES.splice(2)', context);
  const sources = vm.runInContext('RSS_SOURCES', context);
  const fetched = [];
  context.UrlFetchApp.fetchAll = () => { throw new Error('batch failed'); };
  context.UrlFetchApp.fetch = urlOrRequest => {
    const url = typeof urlOrRequest === 'string' ? urlOrRequest : urlOrRequest.url;
    fetched.push(url);
    if (url === sources[1].feedUrl) throw new Error('second source failed');
    return {
      getResponseCode() { return 200; },
      getContentText() { return '<rss version="2.0"><channel/></rss>'; }
    };
  };
  context.XmlService.parse = () => ({
    getRootElement() {
      return {
        getName() { return 'rss'; },
        getNamespace() { return { getURI() { return ''; } }; },
        getChildren() { return [{ getName() { return 'channel'; }, getChildren() { return []; } }]; }
      };
    }
  });
  const results = exports.fetchAllRssFeeds();
  assert.equal(results.length, 2);
  assert.deepEqual(fetched, sources.map(source => source.feedUrl));
  assert.equal(results[0].error, null);
  assert.deepEqual(Array.from(results[0].items), []);
  assert.match(String(results[1].error), /second source failed/);
  assert.deepEqual(Array.from(results[1].items), []);
});

test('RSS failed source does not mutate initialized state, while healthy source continues', () => {
  const { exports, context, propertyValues } = loadStandalone();
  const failed = { id: 'failed', name: 'Failed source' };
  const healthy = { id: 'healthy', name: 'Healthy source' };
  const store = new exports.RssStateStore();
  store.initialize(failed.id, ['old-failed']);
  store.initialize(healthy.id, ['old-healthy']);
  const first = { id: 'new-failed', sourceId: failed.id, publishedAt: new Date(0) };
  const second = { id: 'new-healthy', sourceId: healthy.id, publishedAt: new Date(0) };
  context.__feedResults = [
    { source: failed, items: [first], error: 'HTTP 500' },
    { source: healthy, items: [second], error: null }
  ];
  vm.runInContext('fetchAllRssFeeds = () => globalThis.__feedResults', context);
  const before = propertyValues.get(store.key(failed.id));
  const posted = [];
  exports.processRssNews({ process(item) { posted.push(item.id); return 'posted'; } });
  assert.equal(propertyValues.get(store.key(failed.id)), before);
  assert.deepEqual(posted, ['new-healthy']);
  assert.equal(store.hasProcessed(healthy.id, 'new-healthy'), true);
});

test('email normalization failure is isolated; only successfully processed messages are marked read', () => {
  const { exports, context } = loadStandalone();
  const reads = [];
  const message = (id, failNormalize) => ({
    getId() { return id; },
    getDate() { return new Date('2026-09-25T00:00:00Z'); },
    getFrom() { if (failNormalize) throw new Error('getFrom failed'); return 'sender@example.test'; },
    getSubject() { return id; },
    getPlainBody() { return 'body'; },
    isUnread() { return true; },
    markRead() { reads.push(id); }
  });
  const messages = [message('broken', true), message('success', false),
    message('pipeline-error', false), message('later-success', false)];
  context.GmailApp.search = () => [{ getMessages() { return messages; } }];
  const handled = [];
  exports.processUnreadNewsletters({
    process(item) {
      handled.push(item.id);
      if (item.id === 'pipeline-error') throw new Error('posting failed');
      return 'posted';
    }
  });
  assert.deepEqual(handled, ['success', 'pipeline-error', 'later-success']);
  assert.deepEqual(reads, ['success', 'later-success']);
});

test('XML parser accepts RSS 2.0, Atom and RSS 1.0/RDF but rejects unknown or mismatched roots', () => {
  const { exports, context } = loadStandalone();
  const root = (name, namespace) => ({
    getName() { return name; },
    getNamespace() { return { getURI() { return namespace; } }; },
    getChildren() { return []; }
  });
  context.XmlService.parse = xml => ({ getRootElement() {
    const [name, namespace] = xml.split('|');
    return root(name, namespace || '');
  } });
  const source = { id: 'format', name: 'Format', feedUrl: 'https://example.test/rss' };
  assert.deepEqual(Array.from(exports.parseFeed('rss|', source)), []);
  assert.deepEqual(Array.from(exports.parseFeed('feed|http://www.w3.org/2005/Atom', source)), []);
  assert.deepEqual(Array.from(exports.parseFeed(
    'RDF|http://www.w3.org/1999/02/22-rdf-syntax-ns#', source
  )), []);
  assert.throws(() => exports.parseFeed('html|', source), /XML|feed|RSS|形式|root|ルート/i);
  assert.throws(() => exports.parseFeed('RDF|https://example.test/not-rdf', source),
    /XML|feed|RSS|形式|root|ルート/i);
  assert.throws(() => exports.parseFeed('feed|https://example.test/not-atom', source),
    /XML|feed|RSS|形式|root|ルート/i);
});
