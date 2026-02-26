// =====================================================
// Kiwi 売上進捗ダッシュボード - GAS API
// スプレッドシートの「API」シートからデータを返す
// =====================================================

// -------------------------------------------------------
// AI コンサルタント プロキシ（チャット・学習機能付き）
// Script Properties に ANTHROPIC_API_KEY を設定すること
// -------------------------------------------------------
function doConsult(messagesOrPrompt, store, ctx) {
  try {
    var messages;
    try {
      var parsed = JSON.parse(messagesOrPrompt);
      messages = Array.isArray(parsed) ? parsed : [{ role: 'user', content: messagesOrPrompt }];
    } catch(e) {
      messages = messagesOrPrompt ? [{ role: 'user', content: messagesOrPrompt }] : [];
    }
    if (!messages.length) return out({ error: 'messagesが空です' });

    var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
    if (!apiKey) return out({ error: 'Script PropertiesにANTHROPIC_API_KEYを設定してください' });

    // この店舗の過去の相談履歴を読み込む（学習・蓄積）
    var pastContext = store ? _loadPastContext(store) : '';

    // システムプロンプト: 基本 + 現在のKPIデータ(ctx) + 過去の相談履歴
    var systemPrompt = 'あなたは美容サロン（眉毛・まつ毛パーマ・マツエク専門）の経営コンサルタントです。';
    if (ctx) {
      systemPrompt += '\n\n## 対象店舗の現在のデータ\n' + ctx;
    }
    if (pastContext) {
      systemPrompt += '\n\n## この店舗との過去の相談履歴（参考）\n'
        + pastContext
        + '\n\n過去の相談内容も踏まえて、より的確なアドバイスを提供してください。';
    }
    systemPrompt += '\n\nデータに基づいた具体的で実行可能なアドバイスを日本語で提供してください。';

    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 1200,
        system: systemPrompt,
        messages: messages
      }),
      muteHttpExceptions: true
    });

    var result = JSON.parse(response.getContentText());
    if (result.error) return out({ error: result.error.message || 'Anthropic APIエラー' });

    var text = (result.content && result.content[0]) ? result.content[0].text : '';

    // 会話をスプレッドシートに保存（知識の蓄積）
    if (store) {
      try {
        var lastUserMsg = '';
        for (var i = messages.length - 1; i >= 0; i--) {
          if (messages[i].role === 'user') {
            lastUserMsg = String(messages[i].content).slice(0, 400);
            break;
          }
        }
        _saveLog(store, lastUserMsg, text.slice(0, 1000));
      } catch(logErr) {}
    }

    return out({ text: text });

  } catch(err) {
    return out({ error: String(err) });
  }
}

// この店舗の過去の相談履歴を読み込む（最新5件）
function _loadPastContext(store) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('会話ログ');
    if (!sh || sh.getLastRow() <= 1) return '';

    var allData = sh.getDataRange().getValues();
    var storeLogs = allData.slice(1).filter(function(row) {
      return String(row[1]).trim() === String(store).trim();
    });
    if (!storeLogs.length) return '';

    var recent = storeLogs.slice(-5);
    return recent.map(function(row) {
      var date = row[0] ? Utilities.formatDate(new Date(row[0]), 'Asia/Tokyo', 'M/d') : '';
      var q    = String(row[2] || '').slice(0, 150);
      var a    = String(row[3] || '').slice(0, 200);
      return '[' + date + '] Q: ' + q + '\nA: ' + a;
    }).join('\n\n');
  } catch(e) {
    return '';
  }
}

// 会話ログをスプレッドシートに保存
function _saveLog(store, userMsg, aiMsg) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('会話ログ');
  if (!sh) {
    sh = ss.insertSheet('会話ログ');
    sh.appendRow(['日時', '店舗', 'ユーザー発言', 'AI回答']);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(2, 120);
    sh.setColumnWidth(3, 250);
    sh.setColumnWidth(4, 350);
  }
  sh.appendRow([new Date(), store, userMsg, aiMsg]);
}

function doGet(e) {
  // AIコンサルタントモード: ?action=consult&messages=[...]&store=店舗名&ctx=KPIデータ
  if (e.parameter && e.parameter.action === 'consult') {
    return doConsult(
      e.parameter.messages || e.parameter.prompt || '',
      e.parameter.store || '',
      e.parameter.ctx   || ''
    );
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("API");
    if (!sh) return out({ error: "「API」シートが見つかりません" });

    var lastRow = sh.getLastRow();
    var lastCol = Math.max(sh.getLastColumn(), 42);

    var vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
    var disp = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();

    var dr = vals[0];
    var s0 = String(dr[1] || "");
    var ym = s0.length >= 6
      ? s0.slice(0, 4) + "年" + Number(s0.slice(4, 6)) + "月"
      : "-";

    var stores = [];
    for (var i = 2; i < vals.length; i++) {
      var r  = vals[i];
      var rd = disp[i];
      var nm = String(r[1] || "").trim();
      if (!nm) continue;

      stores.push({
        category_label:      String(r[0] || "").trim(),
        name:                nm,
        kgiGoal:             n(r[2]),
        kgiActual:           n(r[3]),
        kgiForecast:         pct(rd[4],  r[4]),
        unitPriceActual:     n(r[6]),
        unitPriceDiff:       n(r[7]),
        newGuestGoal:        n(r[8]),
        newGuestActual:      n(r[9]),
        newGuestForecast:    pct(rd[10], r[10]),
        repeatGoal:          n(r[11]),
        repeatActual:        n(r[12]),
        repeatForecast:      pct(rd[13], r[13]),
        svGoal:              n(r[27]),
        svActual:            n(r[28]),
        svForecast:          pct(rd[29], r[29]),
        svName:              String(r[30] || "").trim(),
        ticketSalesRate:     pct(rd[31], r[31]),
        ticketUnitPrice:     n(r[32]),
        productSalesRate:    pct(rd[34], r[34]),
        productUnitPrice:    n(r[35]),
        serviceUnitPrice:    n(r[40]),
        nextReservationRate: pct(rd[41], r[41])
      });
    }

    return out({
      dateInfo: {
        yearMonth: ym,
        today:     Number(dr[4]) || new Date().getDate(),
        endDay:    Number(dr[3]) || 31
      },
      stores:        stores,
      bestPractices: []
    });

  } catch (err) {
    return out({ error: String(err) });
  }
}

function pct(dispVal, rawVal) {
  var d = String(dispVal || "").trim();
  if (d.indexOf("%") > -1) {
    var x = parseFloat(d.replace(/[%,\s]/g, ""));
    return isNaN(x) ? 0 : x / 100;
  }
  if (!d || d === "#DIV/0!" || d === "-" || d === "#N/A") return 0;
  if (rawVal !== undefined && rawVal !== null && rawVal !== "") {
    var r = Number(rawVal);
    if (!isNaN(r)) return r;
  }
  var x2 = parseFloat(d.replace(/[,\s]/g, ""));
  return isNaN(x2) ? 0 : x2;
}

function n(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  var s = String(v).replace(/[\s,¥￥\u00A5\uFFE5]/g, "");
  if (!s || s === "#DIV/0!" || s === "-" || s === "#N/A") return 0;
  var x = parseFloat(s);
  return isNaN(x) ? 0 : x;
}

function out(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================
// Slack アラート通知（AI改善案3点付き）
// =====================================================

// Script Properties に SLACK_WEBHOOK_URL を設定してください
// （GASエディタ > 歯車アイコン > スクリプトのプロパティ）
var SLACK_WEBHOOK = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL') || '';

// -------------------------------------------------------
// メイン: Slackにアラートレポートを送信
// GASのトリガーからこの関数を呼び出す
// -------------------------------------------------------
function sendSlackAlerts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('API');
  if (!sh) { _postSlack('⚠️ APIシートが見つかりません'); return; }

  var lastRow = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), 42);
  var vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var disp = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  // 日付情報
  var dr   = vals[0];
  var s0   = String(dr[1] || '');
  var ym   = s0.length >= 6 ? s0.slice(0, 4) + '年' + Number(s0.slice(4, 6)) + '月' : '-';
  var today = Number(dr[4]) || new Date().getDate();

  // アラート店舗を抽出（売上見込み < 100%）
  var alerted = [];
  for (var i = 2; i < vals.length; i++) {
    var r  = vals[i];
    var rd = disp[i];
    var nm = String(r[1] || '').trim();
    if (!nm) continue;

    var fc = pct(rd[4], r[4]);
    if (fc >= 1.0) continue;  // 達成店舗はスキップ

    var alerts = [];
    if (fc < 1.0)                       alerts.push('売上' + Math.round(fc * 100) + '%');
    if (pct(rd[10], r[10]) < 1.0)       alerts.push('新規客未達');
    if (pct(rd[13], r[13]) < 1.0)       alerts.push('再来客未達');
    if (pct(rd[41], r[41]) <= 0.35)     alerts.push('次回予約率' + Math.round(pct(rd[41], r[41]) * 100) + '%');
    if (n(r[7]) < 0)                    alerts.push('客単価低下');
    if (pct(rd[34], r[34]) <= 0.1)      alerts.push('物販率低い');

    alerted.push({
      name:     nm,
      category: String(r[0] || '').trim(),
      svName:   String(r[30] || '').trim(),
      fc:       fc,
      alerts:   alerts,
      kpiText:  '売上' + Math.round(fc * 100) + '% / 新規' + Math.round(pct(rd[10], r[10]) * 100) + '% / 再来' + Math.round(pct(rd[13], r[13]) * 100) + '% / 次回予約' + Math.round(pct(rd[41], r[41]) * 100) + '%'
    });
  }

  // 全店舗OK の場合
  if (alerted.length === 0) {
    _postSlack('✅ *Kiwiアラートレポート | ' + ym + ' ' + today + '日時点*\n全店舗が売上目標を達成しています！');
    return;
  }

  // 売上見込みが低い順にソートして上位5件を処理
  alerted.sort(function(a, b) { return a.fc - b.fc; });
  var targets = alerted.slice(0, 5);

  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');

  // ヘッダー
  var msg = '⚠️ *Kiwiアラートレポート | ' + ym + ' ' + today + '日時点*\n';
  msg += '要注意店舗: *' + alerted.length + '件*（深刻上位' + targets.length + '件を表示）\n';
  msg += '━━━━━━━━━━━━━━━━━━\n';

  for (var j = 0; j < targets.length; j++) {
    var s = targets[j];
    msg += '\n📍 *' + s.name + '*（' + s.category + '）';
    if (s.svName) msg += '　SV: ' + s.svName;
    msg += '\n';
    msg += '　KPI: ' + s.kpiText + '\n';
    msg += '　アラート: ' + s.alerts.join(' · ') + '\n';

    // AI改善案を取得
    if (apiKey) {
      try {
        var suggestions = _getAISuggestions(s, apiKey);
        msg += '　改善案:\n' + suggestions + '\n';
      } catch(e) {
        msg += '　（改善案取得エラー）\n';
      }
    }
    msg += '━━━━━━━━━━━━━━━━━━\n';
  }

  if (alerted.length > 5) {
    msg += '\n他 ' + (alerted.length - 5) + ' 店舗もアラートあり → ダッシュボードで確認\n';
  }
  msg += 'https://tmiyawaki-beep.github.io/kiwi-dashboard/';

  _postSlack(msg);
}

// AI改善案を3点取得（短く・高速）
function _getAISuggestions(store, apiKey) {
  var prompt = store.name + '（' + store.category + '）\n'
    + 'アラート: ' + store.alerts.join('、') + '\n'
    + 'KPI: ' + store.kpiText + '\n\n'
    + '今すぐ実行できる改善アクションを①②③の3点のみ、各1行で簡潔に提案してください。';

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 250,
      messages: [{ role: 'user', content: prompt }]
    }),
    muteHttpExceptions: true
  });
  var result = JSON.parse(res.getContentText());
  if (result.error) return '　（取得失敗）';
  var text = result.content && result.content[0] ? result.content[0].text : '';
  return text.split('\n').filter(function(l) { return l.trim(); })
    .map(function(l) { return '　　' + l; }).join('\n');
}

// Slack にメッセージ送信
function _postSlack(text) {
  UrlFetchApp.fetch(SLACK_WEBHOOK, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: text })
  });
}

// -------------------------------------------------------
// 毎朝8時に自動送信するトリガーを設定（一度だけ実行）
// GASエディタで この関数を選択して「実行」ボタンを押す
// -------------------------------------------------------
function setupDailyTrigger() {
  // 既存のトリガーを削除
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendSlackAlerts') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 毎朝8時に設定
  ScriptApp.newTrigger('sendSlackAlerts')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .inTimezone('Asia/Tokyo')
    .create();
  Logger.log('トリガー設定完了: 毎朝8時にSlack通知');
}
// ================================================================
// Kiwi サロン — 日次情報レポート自動配信
// ================================================================
//
// 【概要】
//   毎朝9時に業界情報を自動収集→Claude でレポート生成
//   →メール送信（t.miyawaki@lime-fit.com）→Slack通知
//
// 【追加先】
//   既存の gas_new_code.gs の末尾に このファイルの内容を貼り付け
//   （または Apps Script で新しいファイル「daily_report」を作成して貼り付け）
//
// 【初回セットアップ手順】
//   1. このコードを Apps Script に貼り付け
//   2. 「保存」後、関数を「testDailyReport」に切り替えて「実行」
//      → Gmail の権限許可を求められるので「許可」を押す
//   3. テストが成功したら「setupDailyReportTrigger」を実行
//      → 毎朝9時（JST）の自動実行が設定される
//
// 【注意】
//   ・ANTHROPIC_API_KEY は Script Properties に設定済みであること
//   ・GAS プロジェクトのタイムゾーンを Asia/Tokyo に設定すること
//     （GAS エディタ > 歯車アイコン > スクリプトのプロパティ > タイムゾーン）
// ================================================================

var _RPT_EMAIL  = 't.miyawaki@lime-fit.com';

// 検索クエリ設定（カテゴリ別）
var _RPT_QUERIES = [
  // ─── 業界トレンド ───
  { c: 'トレンド', q: '眉毛サロン 最新トレンド 人気デザイン 2025' },
  { c: 'トレンド', q: 'まつ毛エクステ 最新技術 新メニュー' },
  { c: 'トレンド', q: 'まつ毛パーマ ナチュラル 人気デザイン' },
  { c: 'トレンド', q: 'ネイル 眉毛 アイラッシュ 複合サロン 新業態' },
  { c: 'トレンド', q: 'アイブロウ スタイリング トレンド 2025' },

  // ─── 競合情報 ───
  { c: '競合', q: 'ロレインブロウ 新店舗 採用 フランチャイズ' },
  { c: '競合', q: "I'm アイブロウサロン フランチャイズ 加盟" },
  { c: '競合', q: 'ホワイトアイ まつ毛サロン 新店舗' },
  { c: '競合', q: 'オーレス まつ毛 サロン' },
  { c: '競合', q: 'マキア まつ毛エクステ サロン 採用' },
  { c: '競合', q: 'ブラン 眉毛サロン 店舗 展開' },
  { c: '競合', q: '眉毛まつ毛専門サロン チェーン 業界 最新' },
  { c: '競合', q: '美容サロン フランチャイズ ランキング 売上' },

  // ─── 商材・技術 ───
  { c: '商材', q: 'パーフェクトラッシュジャパン まつ毛 新商品' },
  { c: '商材', q: '大浴場 まつ毛グルー 接着剤 商材' },
  { c: '商材', q: 'まつ毛エクステ 新素材 持続性 商材 2025' },
  { c: '商材', q: 'アイブロウ ワックス 脱毛 商材 サロン向け' },
  { c: '商材', q: 'まつ毛パーマ 液 ロッド 最新' },

  // ─── SNS・集客 ───
  { c: 'SNS集客', q: 'アイリスト Instagram フォロワー 1000 バズ' },
  { c: 'SNS集客', q: '眉毛サロン Instagram リール 集客 バズ' },
  { c: 'SNS集客', q: 'まつ毛エクステ ビフォーアフター 人気 投稿' },
  { c: 'SNS集客', q: '美容サロン SNS集客 成功事例 低コスト 効果' },
  { c: 'SNS集客', q: '美容師 アイリスト TikTok YouTube バズ集客' },

  // ─── ビジネス戦略 ───
  { c: 'フランチャイズ', q: '美容サロン フランチャイズ 加盟条件 初期費用 2025' },
  { c: 'フランチャイズ', q: '眉毛まつ毛サロン FC 開業 収益モデル' },
  { c: '採用', q: 'アイリスト 採用 方法 成功事例 Instagram 求人' },
  { c: '採用', q: '美容業界 離職率 改善 定着 方法 事例' },
  { c: '採用', q: '美容サロン 福利厚生 充実 人気 事例 2025' },
  { c: '採用', q: '美容師 労働環境 改善 給与 業務委託 雇用' },
  { c: '経営', q: '美容サロン 売上最大化 効率化 低コスト 施策' },
  { c: '経営', q: '美容サロン 客単価 アップ 物販 回数券 成功' },
  { c: '経営', q: '美容サロン リピート率 向上 方法 LTV' },
];

// ================================================================
// メイン関数（トリガーから呼ばれる）
// ================================================================
function generateDailyReport() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    _postSlack('⚠️ 日次レポートエラー: ANTHROPIC_API_KEY が未設定です');
    return;
  }

  try {
    Logger.log('=== 日次レポート開始: ' + new Date().toISOString() + ' ===');

    // 1. ニュース・情報収集
    var newsData = _rpt_collectNews();
    Logger.log('収集完了: ' + newsData.total + '件 / ' + newsData.text.length + '文字');

    // 2. Claude でHTMLレポート生成
    var report = _rpt_generateReport(newsData, apiKey);
    Logger.log('レポート生成完了: ' + report.subject);

    // 3. メール送信
    _rpt_sendEmail(report);
    Logger.log('メール送信完了 → ' + _RPT_EMAIL);

    // 4. Slack通知
    _rpt_notifySlack(report);
    Logger.log('Slack通知完了');

    Logger.log('=== 日次レポート完了 ===');

  } catch (e) {
    var errMsg = '⚠️ 日次レポートエラー: ' + String(e);
    Logger.log(errMsg);
    try { _postSlack(errMsg); } catch(e2) {}
  }
}

// ================================================================
// テスト用関数（初回セットアップ確認に使用）
// ================================================================
function testDailyReport() {
  generateDailyReport();
}

// ================================================================
// 情報収集: Google News RSS を全クエリで取得
// ================================================================
function _rpt_collectNews() {
  var result = { total: 0, byCategory: {}, text: '' };
  var seen = {};  // タイトル重複排除

  for (var i = 0; i < _RPT_QUERIES.length; i++) {
    var cat = _RPT_QUERIES[i].c;
    var q   = _RPT_QUERIES[i].q;

    try {
      var items = _rpt_fetchGNews(q);
      var newItems = items.filter(function(it) {
        if (seen[it.title]) return false;
        seen[it.title] = true;
        return true;
      });

      if (newItems.length > 0) {
        if (!result.byCategory[cat]) result.byCategory[cat] = [];
        result.byCategory[cat] = result.byCategory[cat].concat(newItems);
        result.total += newItems.length;
      }

      Utilities.sleep(200);  // レートリミット対策
    } catch (e) {
      Logger.log('収集エラー[' + q + ']: ' + e);
    }
  }

  // テキスト形式に変換
  var lines = [];
  var cats = Object.keys(result.byCategory);
  for (var ci = 0; ci < cats.length; ci++) {
    var c = cats[ci];
    var its = result.byCategory[c];
    if (!its || its.length === 0) continue;
    lines.push('\n【' + c + '】');
    for (var ii = 0; ii < its.length; ii++) {
      var it = its[ii];
      lines.push('・' + it.title + (it.src ? '（' + it.src + '）' : ''));
      if (it.desc) lines.push('  ' + it.desc);
    }
  }
  result.text = lines.join('\n');

  return result;
}

// ================================================================
// Google News RSS フェッチ & パース
// ================================================================
function _rpt_fetchGNews(query) {
  var url = 'https://news.google.com/rss/search?q='
    + encodeURIComponent(query)
    + '&hl=ja&gl=JP&ceid=JP:ja';

  var res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; KiwiSalonBot/1.0)' }
  });

  if (res.getResponseCode() !== 200) return [];

  var xml = res.getContentText();
  var items = [];
  var rx = /<item>([\s\S]*?)<\/item>/g;
  var m;
  var count = 0;

  while ((m = rx.exec(xml)) !== null && count < 4) {
    var chunk = m[1];

    var title = (chunk.match(/<title><!\[CDATA\[(.*?)\]\]><\/title>/) ||
                 chunk.match(/<title>(.*?)<\/title>/)
                )?.[1] || '';
    var desc  = (chunk.match(/<description><!\[CDATA\[([\s\S]*?)\]\]><\/description>/) ||
                 chunk.match(/<description>([\s\S]*?)<\/description>/)
                )?.[1] || '';
    var src   = chunk.match(/<source[^>]*>(.*?)<\/source>/)?.[1] || '';
    var pub   = chunk.match(/<pubDate>(.*?)<\/pubDate>/)?.[1] || '';

    title = title.replace(/<[^>]+>/g, '').trim();
    desc  = desc.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').substring(0, 250).trim();

    if (title) {
      items.push({ title: title, desc: desc, src: src.trim(), pub: pub.trim() });
      count++;
    }
  }

  return items;
}

// ================================================================
// Claude でHTMLレポートを生成
// ================================================================
function _rpt_generateReport(newsData, apiKey) {
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日(E)');
  var ym    = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  var systemPrompt = [
    'あなたは眉毛・まつ毛・ネイルサロンチェーンの経営コンサルタントです。',
    '【クライアント情報】3ブランド展開: SSIN STUDIO / most eyes / LUMISS',
    '事業内容: 眉毛スタイリング、まつ毛パーマ、まつ毛エクステ、ネイル',
    '直営店とフランチャイズ（SV管理）の混合運営',
    '',
    '【主要競合】ロレインブロウ、I\'m、ホワイトアイ、オーレス、マキア、ブラン',
    '【注目商材】パーフェクトラッシュジャパン、大浴場（まつ毛グルー）',
    '',
    '【レポート優先順位】',
    '1. 集客直結情報（SNS、マーケ施策 ← 最重要）',
    '2. 売上最大化（低コスト・低労力で効果大のもの優先）',
    '3. 採用・定着・福利厚生',
    '4. 競合動向・差別化ヒント',
    '5. ロールモデル施策・成功事例',
    '6. 新商材・技術',
    '7. フランチャイズ・組織情報',
  ].join('\n');

  var userPrompt = [
    '本日（' + today + '）の収集情報（' + newsData.total + '件）:',
    newsData.text.substring(0, 14000),
    '',
    '────────────────────────────────',
    '上記の情報をもとに、以下のJSON形式でレポートを作成してください。',
    '情報が薄いカテゴリはClaudeの業界知識を補完して実用的な内容にしてください。',
    '',
    'JSONフォーマット:',
    '{',
    '  "subject": "メール件名（50文字以内）",',
    '  "slack_summary": "Slack通知用1行サマリー（80文字以内）",',
    '  "html": "メール本文HTML（インラインCSS必須。以下のセクション構成）"',
    '}',
    '',
    '【HTML セクション構成（インラインCSS・モバイル対応）】',
    '① 🏆 今日の注目インサイト TOP3',
    '   - 最も重要な情報を3点。各点に【アクション提案】を1行追加',
    '② 🎯 競合・業界動向',
    '   - ロレインブロウ/I\'m/ホワイトアイ/オーレス/マキア/ブランの動向',
    '   - 差別化・対抗施策のヒント',
    '③ 📱 SNS・Instagram 集客インサイト',
    '   - バズっているアイリスト事例・投稿タイプ',
    '   - 今すぐ真似できる投稿アイデア2〜3点',
    '④ 💰 売上・客単価アップ施策（低コスト・低労力優先）',
    '   - 即実行可能な施策を具体的に',
    '   - 回数券・物販・次回予約の改善ヒント',
    '⑤ 🛍 商材・技術情報',
    '   - パーフェクトラッシュジャパン・大浴場など注目商材',
    '   - 導入検討価値のある新技術・メニュー',
    '⑥ 👥 採用・組織・フランチャイズ',
    '   - 採用成功事例・福利厚生アイデア',
    '   - フランチャイズ動向',
    '⑦ 📝 今週すぐ実行すべきアクション TOP3',
    '   - 優先順位をつけて3点、担当者コメント付き',
    '',
    '重要: HTMLはインラインCSSのみ使用。外部CSSなし。',
    '見出しは適切な色・サイズで視認性高く。モバイルでも読みやすく。',
  ].join('\n');

  var payload = {
    model: 'claude-opus-4-6',
    max_tokens: 6000,
    system: systemPrompt,
    messages: [{ role: 'user', content: userPrompt }]
  };

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var json;
  try {
    json = JSON.parse(res.getContentText());
  } catch (e) {
    throw new Error('Claude API パースエラー: ' + res.getContentText().substring(0, 200));
  }

  if (json.error) throw new Error('Claude API エラー: ' + JSON.stringify(json.error));

  var rawText = (json.content && json.content[0]) ? json.content[0].text : '{}';

  // ```json ... ``` ブロックを除去
  rawText = rawText.replace(/^```json\s*/m, '').replace(/^```\s*/m, '').replace(/\s*```$/m, '').trim();

  var parsed;
  try {
    parsed = JSON.parse(rawText);
  } catch (e) {
    // JSONパース失敗時はテキストをそのまま使用
    Logger.log('JSON parse failed, using raw text. Error: ' + e);
    parsed = {
      subject: today + ' サロン日次レポート',
      slack_summary: '本日の日次レポートを送信しました',
      html: '<p>' + rawText.replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</p>'
    };
  }

  // メール用HTMLにラップ
  var fullHtml = _rpt_wrapEmail(parsed.html || '', parsed.subject || today + ' レポート', today, newsData.total);

  return {
    subject:      parsed.subject || today + ' サロン日次レポート',
    slackSummary: parsed.slack_summary || '本日のレポートを送信しました',
    html:         fullHtml
  };
}

// ================================================================
// メール本文HTMLラッパー（ヘッダー・フッター付き）
// ================================================================
function _rpt_wrapEmail(bodyHtml, subject, today, newsCount) {
  return '<!DOCTYPE html>\n'
    + '<html lang="ja">\n'
    + '<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>' + subject + '</title></head>\n'
    + '<body style="margin:0;padding:16px;background:#eef2f7;font-family:Arial,\'Hiragino Kaku Gothic ProN\',sans-serif">\n'
    + '  <div style="max-width:700px;margin:0 auto">\n'

    // ヘッダー
    + '    <div style="background:linear-gradient(135deg,#0f172a 0%,#1e3a8a 60%,#312e81 100%);'
    + 'border-radius:16px 16px 0 0;padding:24px 28px">\n'
    + '      <p style="color:rgba(255,255,255,0.5);font-size:10px;margin:0 0 6px 0;letter-spacing:.25em;text-transform:uppercase">'
    + 'KIWI SALON CONSULTANT — DAILY INTELLIGENCE REPORT</p>\n'
    + '      <h1 style="color:#ffffff;font-size:17px;margin:0;font-weight:900;line-height:1.5">'
    + subject + '</h1>\n'
    + '      <p style="color:rgba(255,255,255,0.45);font-size:11px;margin:8px 0 0 0">'
    + today + ' ｜ 収集記事 ' + newsCount + '件 ｜ AI分析レポート</p>\n'
    + '    </div>\n'

    // 本文
    + '    <div style="background:#ffffff;padding:28px;border-radius:0 0 16px 16px;'
    + 'box-shadow:0 6px 24px rgba(15,23,42,0.08)">\n'
    + bodyHtml + '\n'
    + '    </div>\n'

    // フッター
    + '    <div style="text-align:center;padding:16px 0">\n'
    + '      <p style="color:#94a3b8;font-size:10px;margin:0">\n'
    + '        Kiwi AI Salon Consultant ｜ 自動配信レポート ｜ '
    + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ' JST\n'
    + '      </p>\n'
    + '    </div>\n'
    + '  </div>\n'
    + '</body>\n</html>';
}

// ================================================================
// メール送信
// ================================================================
function _rpt_sendEmail(report) {
  MailApp.sendEmail({
    to:       _RPT_EMAIL,
    subject:  '📊 ' + report.subject,
    htmlBody: report.html,
    name:     'Kiwi AI Salon Consultant'
  });
}

// ================================================================
// Slack 通知
// ================================================================
function _rpt_notifySlack(report) {
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  var text = [
    '📊 *日次レポート配信完了* (' + now + ')',
    '> ' + report.subject,
    '',
    report.slackSummary,
    '',
    '送信先: ' + _RPT_EMAIL,
    'ダッシュボード: https://tmiyawaki-beep.github.io/kiwi-dashboard/'
  ].join('\n');

  _postSlack(text);
}

// ================================================================
// トリガーセットアップ（初回のみ一度だけ実行）
// GASエディタで「setupDailyReportTrigger」を選択して「実行」
// ================================================================
function setupDailyReportTrigger() {
  // 既存の日次レポートトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'generateDailyReport') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 毎朝9時（プロジェクトタイムゾーン = Asia/Tokyo）に実行
  ScriptApp.newTrigger('generateDailyReport')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log('✅ トリガー設定完了: 毎朝9時に generateDailyReport を実行');
  Logger.log('   ※ GASプロジェクト設定のタイムゾーンが Asia/Tokyo であることを確認してください');

  // Slack に設定完了通知
  try {
    _postSlack('✅ 日次レポートトリガー設定完了\n毎朝9時に ' + _RPT_EMAIL + ' へ自動配信します');
  } catch(e) {}
}
