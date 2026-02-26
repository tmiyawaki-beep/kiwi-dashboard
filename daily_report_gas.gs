// ================================================================
// Kiwi サロン — 日次情報レポート自動配信（スタンドアロン版）
// ================================================================
//
// 【このファイルだけで完結します】ダッシュボードGASとは独立して動作
//
// 【新しいスプレッドシートでのセットアップ】
//   1. Google スプレッドシートを新規作成
//   2. 「拡張機能」→「Apps Script」を開く
//   3. このファイルの内容を全て貼り付けて保存（Ctrl+S）
//   4. 「⚙️ プロジェクトの設定」→「スクリプトのプロパティ」に以下を追加:
//        ANTHROPIC_API_KEY  = （Anthropicのキー）
//        SLACK_WEBHOOK_URL  = https://hooks.slack.com/services/...
//   5. タイムゾーンを Asia/Tokyo に設定
//   6. 関数「testDailyReport」を実行 → Gmail権限を許可
//   7. 関数「setupDailyReportTrigger」を実行 → 毎朝9時の自動実行が設定される
//
// 【実行後の動作】
//   ・毎朝9時: 業界情報を収集 → Claude がレポート生成
//             → t.miyawaki@lime-fit.com にメール送信
//             → Slack に完了通知
//   ・送信済みレポートはスプレッドシートの「レポート履歴」シートに蓄積
// ================================================================

// ─── 設定 ───────────────────────────────────────────────────────
var RPT_TO    = 't.miyawaki@lime-fit.com';
var RPT_FROM  = 'Kiwi AI Salon Consultant';

// 検索クエリ（カテゴリ別）
var RPT_QUERIES = [
  // トレンド
  { c: 'トレンド', q: '眉毛サロン 最新トレンド 人気デザイン 2025' },
  { c: 'トレンド', q: 'まつ毛エクステ 最新技術 新メニュー' },
  { c: 'トレンド', q: 'まつ毛パーマ ナチュラル 人気デザイン' },
  { c: 'トレンド', q: 'ネイル 眉毛 アイラッシュ 複合サロン 新業態' },
  { c: 'トレンド', q: 'アイブロウ スタイリング トレンド 2025' },
  // 競合
  { c: '競合', q: 'ロレインブロウ 新店舗 採用 フランチャイズ' },
  { c: '競合', q: "I'm アイブロウサロン フランチャイズ 加盟" },
  { c: '競合', q: 'ホワイトアイ まつ毛サロン 新店舗' },
  { c: '競合', q: 'オーレス まつ毛 サロン' },
  { c: '競合', q: 'マキア まつ毛エクステ サロン 採用' },
  { c: '競合', q: 'ブラン 眉毛サロン 店舗 展開' },
  { c: '競合', q: '眉毛まつ毛専門サロン チェーン 最新動向' },
  { c: '競合', q: '美容サロン フランチャイズ ランキング 売上' },
  // 商材・技術
  { c: '商材', q: 'パーフェクトラッシュジャパン まつ毛 新商品' },
  { c: '商材', q: '大浴場 まつ毛グルー 接着剤 商材' },
  { c: '商材', q: 'まつ毛エクステ 新素材 持続性 商材 2025' },
  { c: '商材', q: 'アイブロウ ワックス 脱毛 商材 サロン向け' },
  { c: '商材', q: 'まつ毛パーマ 液 ロッド 最新' },
  // SNS・集客
  { c: 'SNS集客', q: 'アイリスト Instagram フォロワー バズ 集客' },
  { c: 'SNS集客', q: '眉毛サロン Instagram リール 集客 バズ' },
  { c: 'SNS集客', q: 'まつ毛エクステ ビフォーアフター 人気 投稿' },
  { c: 'SNS集客', q: '美容サロン SNS集客 成功事例 低コスト 効果' },
  { c: 'SNS集客', q: '美容師 アイリスト TikTok YouTube バズ集客' },
  // フランチャイズ
  { c: 'フランチャイズ', q: '美容サロン フランチャイズ 加盟条件 初期費用 2025' },
  { c: 'フランチャイズ', q: '眉毛まつ毛サロン FC 開業 収益モデル' },
  // 採用・組織
  { c: '採用', q: 'アイリスト 採用 方法 成功事例 Instagram 求人' },
  { c: '採用', q: '美容業界 離職率 改善 定着 方法 事例' },
  { c: '採用', q: '美容サロン 福利厚生 充実 人気 事例 2025' },
  { c: '採用', q: '美容師 労働環境 改善 給与 業務委託 雇用' },
  // 経営・売上
  { c: '経営', q: '美容サロン 売上最大化 効率化 低コスト 施策' },
  { c: '経営', q: '美容サロン 客単価 アップ 物販 回数券 成功' },
  { c: '経営', q: '美容サロン リピート率 向上 方法 LTV 次回予約' },
];

// ================================================================
// メイン関数（毎朝9時にトリガーから呼ばれる）
// ================================================================
function generateDailyReport() {
  var props  = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('ANTHROPIC_API_KEY');
  var slack  = props.getProperty('SLACK_WEBHOOK_URL') || '';

  if (!apiKey) {
    rpt_postSlack(slack, '⚠️ 日次レポートエラー: ANTHROPIC_API_KEY が未設定です');
    return;
  }

  try {
    Logger.log('=== 日次レポート開始: ' + new Date().toISOString() + ' ===');

    // 1. ニュース・情報収集
    var newsData = rpt_collectNews();
    Logger.log('収集完了: ' + newsData.total + '件');

    // 2. Claude でHTMLレポート生成
    var report = rpt_generateReport(newsData, apiKey);
    Logger.log('レポート生成完了: ' + report.subject);

    // 3. メール送信
    rpt_sendEmail(report);
    Logger.log('メール送信完了 → ' + RPT_TO);

    // 4. スプレッドシートに履歴保存
    rpt_saveHistory(report);
    Logger.log('履歴保存完了');

    // 5. Slack通知
    rpt_notifySlack(slack, report);
    Logger.log('Slack通知完了');

    Logger.log('=== 日次レポート完了 ===');

  } catch (e) {
    var errMsg = '⚠️ 日次レポートエラー: ' + String(e);
    Logger.log(errMsg);
    try { rpt_postSlack(slack, errMsg); } catch (e2) {}
  }
}

// テスト用（初回セットアップ確認）
function testDailyReport() {
  generateDailyReport();
}

// ================================================================
// ニュース収集: Google News RSS を全クエリで取得
// ================================================================
function rpt_collectNews() {
  var result = { total: 0, byCategory: {}, text: '' };
  var seen   = {};

  for (var i = 0; i < RPT_QUERIES.length; i++) {
    var cat = RPT_QUERIES[i].c;
    var q   = RPT_QUERIES[i].q;

    try {
      var items = rpt_fetchGNews(q);
      var fresh = items.filter(function(it) {
        if (seen[it.title]) return false;
        seen[it.title] = true;
        return true;
      });

      if (fresh.length > 0) {
        if (!result.byCategory[cat]) result.byCategory[cat] = [];
        result.byCategory[cat] = result.byCategory[cat].concat(fresh);
        result.total += fresh.length;
      }
      Utilities.sleep(200);
    } catch (e) {
      Logger.log('収集エラー[' + q + ']: ' + e);
    }
  }

  // テキスト変換
  var lines = [];
  var cats  = Object.keys(result.byCategory);
  for (var ci = 0; ci < cats.length; ci++) {
    var c  = cats[ci];
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
function rpt_fetchGNews(query) {
  var url = 'https://news.google.com/rss/search?q='
    + encodeURIComponent(query) + '&hl=ja&gl=JP&ceid=JP:ja';

  var res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; KiwiSalonBot/1.0)' }
  });
  if (res.getResponseCode() !== 200) return [];

  var xml   = res.getContentText();
  var items = [];
  var rx    = /<item>([\s\S]*?)<\/item>/g;
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

    title = title.replace(/<[^>]+>/g, '').trim();
    desc  = desc.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').substring(0, 250).trim();

    if (title) {
      items.push({ title: title, desc: desc, src: src.trim() });
      count++;
    }
  }
  return items;
}

// ================================================================
// Claude でHTMLレポート生成
// ================================================================
function rpt_generateReport(newsData, apiKey) {
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日(E)');

  var systemPrompt = [
    'あなたは眉毛・まつ毛・ネイルサロンチェーンの経営コンサルタントです。',
    '【クライアント情報】3ブランド: SSIN STUDIO / most eyes / LUMISS',
    '事業: 眉毛スタイリング、まつ毛パーマ、まつ毛エクステ、ネイル',
    '直営店とフランチャイズ（SV管理）の混合運営',
    '',
    '【主要競合】ロレインブロウ、I\'m、ホワイトアイ、オーレス、マキア、ブラン',
    '【注目商材】パーフェクトラッシュジャパン、大浴場（まつ毛グルー）',
    '',
    '【レポート優先順位】',
    '1. 集客直結情報（SNS・マーケ施策 ← 最重要）',
    '2. 売上最大化（低コスト・低労力で効果大を優先）',
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
    '以下のJSON形式でレポートを作成してください。',
    '情報が薄いカテゴリはClaudeの業界知識で補完して実用的な内容にしてください。',
    '',
    '{',
    '  "subject": "メール件名（50文字以内）",',
    '  "slack_summary": "Slack通知用1行サマリー（80文字以内）",',
    '  "html": "メール本文HTML（インラインCSSのみ使用。以下のセクション構成）"',
    '}',
    '',
    '【HTML セクション構成】',
    '① 🏆 今日の注目インサイト TOP3',
    '   最重要情報を3点。各点に【→ アクション提案】を1行追加',
    '② 🎯 競合・業界動向',
    '   ロレインブロウ/I\'m/ホワイトアイ/オーレス/マキア/ブランの動向',
    '   差別化・対抗施策のヒント',
    '③ 📱 SNS・Instagram 集客インサイト',
    '   バズっているアイリスト・投稿タイプの傾向',
    '   今すぐ真似できる投稿アイデア2〜3点（具体的に）',
    '④ 💰 売上・客単価アップ施策（低コスト・低労力優先）',
    '   即実行可能な施策を具体的に',
    '   回数券・物販・次回予約の改善ヒント',
    '⑤ 🛍 商材・技術情報',
    '   パーフェクトラッシュジャパン・大浴場など注目商材の最新情報',
    '   導入検討価値のある新技術・メニュー',
    '⑥ 👥 採用・組織・フランチャイズ',
    '   採用成功事例・福利厚生アイデア（具体例付き）',
    '   フランチャイズ動向',
    '⑦ 📝 今週すぐ実行すべきアクション TOP3',
    '   優先順位付きで3点。担当者レベルで実行できる粒度で',
    '',
    '重要: HTMLはインラインCSSのみ。外部CSS不可。',
    '見出しは色・サイズで視認性を高く。モバイルでも読みやすく。',
  ].join('\n');

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:      'post',
    contentType: 'application/json',
    headers: {
      'x-api-key':          apiKey,
      'anthropic-version':  '2023-06-01'
    },
    payload: JSON.stringify({
      model:      'claude-opus-4-6',
      max_tokens: 6000,
      system:     systemPrompt,
      messages:   [{ role: 'user', content: userPrompt }]
    }),
    muteHttpExceptions: true
  });

  var json;
  try { json = JSON.parse(res.getContentText()); }
  catch (e) { throw new Error('Claude APIパースエラー: ' + res.getContentText().substring(0, 300)); }

  if (json.error) throw new Error('Claude APIエラー: ' + JSON.stringify(json.error));

  var rawText = (json.content && json.content[0]) ? json.content[0].text : '{}';
  rawText = rawText.replace(/^```json\s*/m, '').replace(/^```\s*/m, '').replace(/\s*```$/m, '').trim();

  var parsed;
  try {
    parsed = JSON.parse(rawText);
  } catch (e) {
    Logger.log('JSONパース失敗、フォールバック使用');
    parsed = {
      subject:       today + ' サロン日次レポート',
      slack_summary: '本日のレポートを送信しました',
      html:          '<p style="font-size:14px;line-height:1.8">' + rawText.replace(/\n/g, '<br>') + '</p>'
    };
  }

  var fullHtml = rpt_wrapEmail(parsed.html || '', parsed.subject || today + ' レポート', today, newsData.total);

  return {
    subject:      parsed.subject      || today + ' サロン日次レポート',
    slackSummary: parsed.slack_summary || '本日のレポートを送信しました',
    bodyHtml:     parsed.html          || '',
    html:         fullHtml
  };
}

// ================================================================
// メール本文HTML ラッパー
// ================================================================
function rpt_wrapEmail(bodyHtml, subject, today, newsCount) {
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  return '<!DOCTYPE html>\n<html lang="ja">\n'
    + '<head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>' + subject + '</title></head>\n'
    + '<body style="margin:0;padding:16px;background:#eef2f7;'
    + 'font-family:Arial,\'Hiragino Kaku Gothic ProN\',sans-serif">\n'
    + '<div style="max-width:700px;margin:0 auto">\n'

    // ヘッダー
    + '<div style="background:linear-gradient(135deg,#0f172a 0%,#1e3a8a 60%,#312e81 100%);'
    + 'border-radius:16px 16px 0 0;padding:24px 28px">\n'
    + '<p style="color:rgba(255,255,255,0.5);font-size:10px;margin:0 0 6px;'
    + 'letter-spacing:.25em;text-transform:uppercase">'
    + 'KIWI SALON CONSULTANT — DAILY INTELLIGENCE REPORT</p>\n'
    + '<h1 style="color:#fff;font-size:17px;margin:0;font-weight:900;line-height:1.5">'
    + subject + '</h1>\n'
    + '<p style="color:rgba(255,255,255,0.45);font-size:11px;margin:8px 0 0">'
    + today + '｜収集記事 ' + newsCount + '件｜AI分析レポート</p>\n'
    + '</div>\n'

    // 本文
    + '<div style="background:#fff;padding:28px;border-radius:0 0 16px 16px;'
    + 'box-shadow:0 6px 24px rgba(15,23,42,0.08)">\n'
    + bodyHtml + '\n'
    + '</div>\n'

    // フッター
    + '<div style="text-align:center;padding:16px 0">\n'
    + '<p style="color:#94a3b8;font-size:10px;margin:0">'
    + 'Kiwi AI Salon Consultant｜自動配信｜' + now + ' JST</p>\n'
    + '</div>\n'
    + '</div>\n</body>\n</html>';
}

// ================================================================
// メール送信
// ================================================================
function rpt_sendEmail(report) {
  MailApp.sendEmail({
    to:       RPT_TO,
    subject:  '📊 ' + report.subject,
    htmlBody: report.html,
    name:     RPT_FROM
  });
}

// ================================================================
// スプレッドシートにレポート履歴を保存
// ================================================================
function rpt_saveHistory(report) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('レポート履歴');

  if (!sh) {
    sh = ss.insertSheet('レポート履歴');
    sh.appendRow(['送信日時', '件名', 'サマリー', '送信先']);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 280);
    sh.setColumnWidth(3, 320);
    sh.setColumnWidth(4, 160);
    sh.getRange('1:1').setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold');
  }

  sh.appendRow([
    new Date(),
    report.subject,
    report.slackSummary,
    RPT_TO
  ]);
}

// ================================================================
// Slack 通知
// ================================================================
function rpt_notifySlack(webhookUrl, report) {
  if (!webhookUrl) return;
  var now  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  var text = [
    '📊 *日次レポート配信完了* (' + now + ')',
    '> ' + report.subject,
    '',
    report.slackSummary,
    '',
    '送信先: ' + RPT_TO
  ].join('\n');
  rpt_postSlack(webhookUrl, text);
}

// ================================================================
// Slack 送信（汎用）
// ================================================================
function rpt_postSlack(webhookUrl, text) {
  if (!webhookUrl) return;
  UrlFetchApp.fetch(webhookUrl, {
    method:      'post',
    contentType: 'application/json',
    payload:     JSON.stringify({ text: text })
  });
}

// ================================================================
// トリガーセットアップ（初回のみ実行）
// ================================================================
function setupDailyReportTrigger() {
  // 既存の同名トリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'generateDailyReport') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 毎朝9時（Asia/Tokyo）
  ScriptApp.newTrigger('generateDailyReport')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log('✅ トリガー設定完了: 毎朝9時に generateDailyReport を実行');

  var slack = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL') || '';
  try {
    rpt_postSlack(slack,
      '✅ 日次レポート トリガー設定完了\n毎朝9時に ' + RPT_TO + ' へ自動配信します');
  } catch (e) {}
}
