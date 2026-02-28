// ================================================================
// Kiwi サロン -- 日次情報レポート自動配信（AI自動学習版）
// ================================================================
//
// 【毎日異なる内容を配信するための3つの仕組み】
//   1. 過去14日分のレポート履歴をClaudeに渡して繰り返し防止
//   2. 曜日別テーマローテーション（月=売上、火=競合...）
//   3. 日付フィルター付きRSSで最新ニュースのみ取得
//
// 【新しいスプレッドシートでのセットアップ】
//   1. Google スプレッドシートを新規作成
//   2. 「拡張機能」→「Apps Script」を開く
//   3. このファイルの内容を全て貼り付けて保存（Ctrl+S）
//   4. 「プロジェクトの設定」→「スクリプトのプロパティ」に以下を追加:
//        ANTHROPIC_API_KEY  = （Anthropicのキー）
//        SLACK_WEBHOOK_URL  = https://hooks.slack.com/services/...
//   5. タイムゾーンを Asia/Tokyo に設定
//   6. 関数「testDailyReport」を実行 → Gmail権限を許可
//   7. 関数「setupDailyReportTrigger」を実行 → 毎朝8時自動実行が設定される
// ================================================================

// --- 設定 ------------------------------------------------------------
var RPT_TO   = 't.miyawaki@lime-fit.com';
var RPT_FROM = 'Kiwi AI Salon Consultant';

// 曜日別テーマ（0=日 〜 6=土）
// 毎日違う視点でClaudeが分析するため、繰り返しを防ぐ
var RPT_DAILY_FOCUS = [
  { theme: 'SNS集客・Instagram戦略の日',    focus: 'バズ投稿・リール・ビフォーアフター・フォロワー獲得・インフルエンサー活用' },
  { theme: '売上最大化・客単価アップの日',   focus: '回数券販売・物販推進・次回予約・単価アップ・LTV改善策' },
  { theme: '競合分析・差別化戦略の日',       focus: 'ロレインブロウ・I\'m・ホワイトアイ・オーレス・マキア・ブランの最新動向と対抗策' },
  { theme: '採用・スタッフ定着の日',         focus: '採用成功事例・福利厚生・離職防止・給与モデル・職場環境改善' },
  { theme: '商材・技術革新の日',             focus: 'パーフェクトラッシュ・大浴場グルー・新素材・施術技術・新メニュー開発' },
  { theme: 'フランチャイズ・多店舗経営の日', focus: 'FC加盟動向・SV管理・多店舗展開・収益モデル・本部サポート強化' },
  { theme: 'トレンド・市場動向の日',         focus: '最新デザイン・新業態・顧客ニーズ変化・市場規模・海外トレンド' }
];

// クエリプール（カテゴリ + 優先曜日）
// ALL=全曜日、または曜日番号配列で指定
var RPT_ALL_QUERIES = [
  // --- トレンド ---
  { c: 'トレンド', q: '眉毛サロン 最新トレンド 人気デザイン',           days: 'ALL' },
  { c: 'トレンド', q: 'まつ毛エクステ 最新技術 新メニュー',             days: 'ALL' },
  { c: 'トレンド', q: 'まつ毛パーマ ナチュラル 人気',                   days: [0,3,6] },
  { c: 'トレンド', q: 'アイブロウ スタイリング 新デザイン',             days: [0,6] },
  { c: 'トレンド', q: 'ネイル 眉毛 アイラッシュ 複合サロン 新業態',    days: [6] },
  { c: 'トレンド', q: '美容サロン 流行 顧客ニーズ 新規客',             days: [1,6] },
  // --- 競合 ---
  { c: '競合',     q: 'ロレインブロウ フランチャイズ 新店舗 採用',      days: [2,6] },
  { c: '競合',     q: "I'm アイブロウサロン FC 加盟 拡大",              days: [2] },
  { c: '競合',     q: 'ホワイトアイ まつ毛サロン 動向',                 days: [2] },
  { c: '競合',     q: 'オーレス まつ毛 サロン 展開',                    days: [2] },
  { c: '競合',     q: 'マキア まつ毛 サロン 採用 展開',                 days: [2] },
  { c: '競合',     q: 'ブラン 眉毛サロン 店舗 展開',                    days: [2] },
  { c: '競合',     q: '眉毛まつ毛専門サロン チェーン 最新動向',         days: [2,6] },
  { c: '競合',     q: '美容サロン フランチャイズ ランキング 売上',       days: [2,5] },
  // --- 商材・技術 ---
  { c: '商材',     q: 'パーフェクトラッシュジャパン まつ毛 新商品',      days: 'ALL' },
  { c: '商材',     q: '大浴場 まつ毛グルー 接着剤 商材',                days: 'ALL' },
  { c: '商材',     q: 'まつ毛エクステ 新素材 持続 商材',                days: [3,4] },
  { c: '商材',     q: 'アイブロウ ワックス 商材 新商品',                days: [4] },
  { c: '商材',     q: 'まつ毛パーマ 液 ロッド 最新',                    days: [3,4] },
  { c: '商材',     q: 'サロン 美容液 ケア商材 ホームケア',              days: [1,4] },
  // --- SNS集客 ---
  { c: 'SNS集客',  q: 'アイリスト Instagram フォロワー バズ 集客',      days: 'ALL' },
  { c: 'SNS集客',  q: '眉毛サロン Instagram リール 集客',               days: [0,1] },
  { c: 'SNS集客',  q: 'まつ毛エクステ ビフォーアフター 人気 投稿',      days: [0,1] },
  { c: 'SNS集客',  q: '美容サロン SNS集客 成功事例 低コスト',           days: [0,1] },
  { c: 'SNS集客',  q: '美容師 アイリスト TikTok バズ 集客',             days: [0] },
  { c: 'SNS集客',  q: 'サロン リール ショート動画 バズ 集客施策',       days: [0,6] },
  // --- フランチャイズ ---
  { c: 'FC',       q: '美容サロン フランチャイズ 加盟条件 収益',         days: [5] },
  { c: 'FC',       q: '眉毛まつ毛サロン FC 開業 収益モデル',            days: [5] },
  { c: 'FC',       q: 'サロン フランチャイズ 本部 サポート 成功',        days: [5] },
  { c: 'FC',       q: '美容 FC 加盟店 オーナー 口コミ 評判',            days: [5] },
  // --- 採用・組織 ---
  { c: '採用',     q: 'アイリスト 採用 方法 成功事例 Instagram',         days: 'ALL' },
  { c: '採用',     q: '美容業界 離職率 改善 定着 方法 事例',            days: [3] },
  { c: '採用',     q: '美容サロン 福利厚生 充実 人気 事例',             days: [3] },
  { c: '採用',     q: '美容師 労働環境 改善 給与 業務委託',             days: [3] },
  { c: '採用',     q: 'サロン 採用 SNS Instagram 求人 成功',            days: [3,0] },
  { c: '採用',     q: '美容師 就職 新卒 Z世代 価値観 職場選び',         days: [3] },
  // --- 経営・売上 ---
  { c: '経営',     q: '美容サロン 売上最大化 効率化 低コスト 施策',      days: 'ALL' },
  { c: '経営',     q: '美容サロン 客単価 アップ 物販 回数券',            days: [1] },
  { c: '経営',     q: '美容サロン リピート率 向上 LTV 次回予約',         days: [1] },
  { c: '経営',     q: 'サロン 経営 集客 コスト削減 成功事例',            days: [1,5] },
  { c: '経営',     q: '美容サロン ロールモデル 繁盛店 成功事例',         days: [1,6] },
  { c: '経営',     q: 'サロン DX 予約 管理 効率化 システム',             days: [5] }
];

// ================================================================
// 過去レポート履歴の読み込み（繰り返し防止に使用）
// ================================================================
function rpt_loadPastReports() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('レポート履歴');
    if (!sh || sh.getLastRow() < 2) return null;

    var lastRow = sh.getLastRow();
    // 直近14日分を読む（ヘッダー行を除く）
    var startRow = Math.max(2, lastRow - 13);
    var numRows  = lastRow - startRow + 1;
    var rows = sh.getRange(startRow, 1, numRows, 5).getValues();

    var lines = ['【過去レポート履歴（重複禁止リスト）】'];
    for (var i = rows.length - 1; i >= 0; i--) {
      if (!rows[i][0]) continue;
      var d   = Utilities.formatDate(new Date(rows[i][0]), 'Asia/Tokyo', 'MM/dd(E)');
      var sub = rows[i][1] || '';
      var sum = rows[i][2] || '';
      var thm = rows[i][4] || '';
      lines.push(d + ' [' + thm + '] ' + sub + ' | ' + sum);
    }
    return lines.join('\n');
  } catch (e) {
    Logger.log('履歴読み込みエラー: ' + e);
    return null;
  }
}

// ================================================================
// メイン関数（毎朝8時にトリガーから呼ばれる）
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
    var now = new Date();
    Logger.log('=== 日次レポート開始: ' + now.toISOString() + ' ===');

    // 曜日テーマ取得（0=日 〜 6=土）
    var dow   = now.getDay();
    var focus = RPT_DAILY_FOCUS[dow];
    Logger.log('本日のテーマ: ' + focus.theme);

    // 過去レポート読み込み（Claude の繰り返し防止に使用）
    var pastContext = rpt_loadPastReports();
    Logger.log('過去履歴: ' + (pastContext ? '読み込み完了' : 'なし（初回）'));

    // 2日前の日付でRSSフィルター → 最新ニュースのみ取得
    var filterDate = new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000);
    var dateFilter = Utilities.formatDate(filterDate, 'Asia/Tokyo', 'yyyy-MM-dd');
    Logger.log('日付フィルター: after:' + dateFilter);

    // ニュース収集
    var newsData = rpt_collectNews(dow, dateFilter);
    Logger.log('収集完了: ' + newsData.total + '件');

    // Claude でHTMLレポート生成
    var report = rpt_generateReport(newsData, apiKey, pastContext, focus);
    Logger.log('レポート生成完了: ' + report.subject);

    // メール送信
    rpt_sendEmail(report);
    Logger.log('メール送信完了 → ' + RPT_TO);

    // 履歴保存（テーマも記録）
    rpt_saveHistory(report, focus.theme);
    Logger.log('履歴保存完了');

    // Slack通知
    rpt_notifySlack(slack, report);
    Logger.log('Slack通知完了');

    Logger.log('=== 日次レポート完了 ===');

  } catch (e) {
    var errMsg = '⚠️ 日次レポートエラー: ' + String(e);
    Logger.log(errMsg);
    try { rpt_postSlack(slack, errMsg); } catch (e2) {}
  }
}

// テスト用
function testDailyReport() {
  generateDailyReport();
}

// ================================================================
// ニュース収集（曜日別クエリ + 日付フィルター）
// ================================================================
function rpt_collectNews(dow, dateFilter) {
  var result = { total: 0, byCategory: {}, text: '' };
  var seen   = {};

  // 当日の曜日に対応するクエリを抽出
  var queries = RPT_ALL_QUERIES.filter(function(q) {
    if (q.days === 'ALL') return true;
    return q.days.indexOf(dow) !== -1;
  });

  Logger.log('使用クエリ数: ' + queries.length + '件（曜日' + dow + '）');

  for (var i = 0; i < queries.length; i++) {
    var cat = queries[i].c;
    var q   = queries[i].q;

    try {
      var items = rpt_fetchGNews(q, dateFilter);
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

  // フィルター後ニュースが少ない場合、フィルターなしで補完
  if (result.total < 10) {
    Logger.log('記事不足（' + result.total + '件）、日付フィルターなしで補完');
    for (var j = 0; j < queries.length && result.total < 30; j++) {
      var cat2 = queries[j].c;
      var q2   = queries[j].q;
      try {
        var items2 = rpt_fetchGNews(q2, null);
        var fresh2 = items2.filter(function(it) {
          if (seen[it.title]) return false;
          seen[it.title] = true;
          return true;
        });
        if (fresh2.length > 0) {
          if (!result.byCategory[cat2]) result.byCategory[cat2] = [];
          result.byCategory[cat2] = result.byCategory[cat2].concat(fresh2);
          result.total += fresh2.length;
        }
        Utilities.sleep(200);
      } catch (e) {
        Logger.log('補完エラー[' + q2 + ']: ' + e);
      }
    }
  }

  // テキスト変換
  var lines = [];
  var cats  = Object.keys(result.byCategory);
  for (var ci = 0; ci < cats.length; ci++) {
    var c   = cats[ci];
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
// Google News RSS フェッチ（日付フィルター対応）
// ================================================================
function rpt_fetchGNews(query, dateFilter) {
  // dateFilter = 'yyyy-MM-dd' を指定すると after:YYYY-MM-DD を付加して最新記事のみ取得
  var q = dateFilter ? query + ' after:' + dateFilter : query;
  var url = 'https://news.google.com/rss/search?q='
    + encodeURIComponent(q) + '&hl=ja&gl=JP&ceid=JP:ja';

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

  while ((m = rx.exec(xml)) !== null && count < 5) {
    var chunk = m[1];
    var title = (chunk.match(/<title><!\[CDATA\[(.*?)\]\]><\/title>/) ||
                 chunk.match(/<title>(.*?)<\/title>/)
                 )?.[1] || '';
    var desc  = (chunk.match(/<description><!\[CDATA\[([\s\S]*?)\]\]><\/description>/) ||
                 chunk.match(/<description>([\s\S]*?)<\/description>/)
                 )?.[1] || '';
    var src   = chunk.match(/<source[^>]*>(.*?)<\/source>/)?.[1] || '';

    title = title.replace(/<[^>]+>/g, '').trim();
    desc  = desc.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').substring(0, 200).trim();

    if (title) {
      items.push({ title: title, desc: desc, src: src.trim() });
      count++;
    }
  }
  return items;
}

// ================================================================
// Claude でHTMLレポート生成（過去履歴 + 曜日テーマ対応）
// ================================================================
function rpt_generateReport(newsData, apiKey, pastContext, focus) {
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日(E)');

  // ---- システムプロンプト ----
  var systemLines = [
    'あなたはKiwiサロングループ専属のビジネスインテリジェンスアナリストです。',
    '3ブランド（SSIN STUDIO / most eyes / LUMISS）を運営する眉毛・まつ毛・ネイルサロンチェーンのオーナーに',
    '毎日新鮮かつ実用的なインサイトを届けることがあなたの使命です。',
    '',
    '【クライアント事業概要】',
    '  事業: 眉毛スタイリング、まつ毛パーマ、まつ毛エクステ、ネイル',
    '  運営: 直営店 + フランチャイズ（SV管理）の混合モデル',
    '  主要競合: ロレインブロウ、I\'m、ホワイトアイ、オーレス、マキア、ブラン',
    '  注目商材: パーフェクトラッシュジャパン、大浴場（まつ毛グルー）',
    '',
    '【レポート品質原則】',
    '  - 毎日「新しい発見」を届ける。同じ内容を繰り返さない',
    '  - 情報が薄い日でもClaudeの業界知識で補完して実用的な内容にする',
    '  - 抽象論より具体的アクション提案（明日から実行できる粒度で）',
    '  - 優先度: 集客直結 > 売上最大化（低コスト低労力） > 採用定着 > 競合 > 商材 > FC',
  ];

  // 過去履歴がある場合は追加
  if (pastContext) {
    systemLines.push('');
    systemLines.push('【絶対に繰り返してはいけない過去レポート内容】');
    systemLines.push(pastContext);
    systemLines.push('');
    systemLines.push('↑ 上記の過去レポートで既に取り上げたトピック・施策・競合情報は今日のレポートに含めないこと。');
    systemLines.push('  新しい角度・未カバーの情報・より深掘りした視点で構成すること。');
  }

  var systemPrompt = systemLines.join('\n');

  // ---- ユーザープロンプト ----
  var userLines = [
    '=== 本日（' + today + '）のレポート生成 ===',
    '',
    '【本日の特集テーマ】' + focus.theme,
    '【重点キーワード】' + focus.focus,
    '',
    '【収集情報（' + newsData.total + '件）】',
    newsData.text.substring(0, 12000),
    '',
    '────────────────────────────────',
    '以下のJSON形式でレポートを作成してください。',
    '収集情報が薄い場合はClaudeの知識で補完し、必ず実用的なインサイトを盛り込んでください。',
    '',
    '{',
    '  "subject": "件名（例: 【月曜】回数券転換率30%超の施策3選 ← 本日テーマを反映した具体的な件名、50文字以内）",',
    '  "slack_summary": "Slack用1行サマリー（本日の最重要発見を1文で、80文字以内）",',
    '  "html": "メール本文HTML（インラインCSSのみ。下記セクション構成）"',
    '}',
    '',
    '【HTMLセクション構成（本日テーマを①に反映）】',
    '① 🏆 今日の注目インサイト TOP3（本日テーマ「' + focus.theme + '」に関連した新発見）',
    '   各インサイトに【→ 今すぐできるアクション】を1行追加',
    '② 🎯 競合・業界動向',
    '   ロレインブロウ / I\'m / ホワイトアイ / オーレス / マキア / ブランの動向',
    '   各社への差別化・対抗施策ヒント',
    '③ 📱 SNS・Instagram集客インサイト',
    '   バズっている投稿タイプの傾向と真似できるアイデア2〜3点（具体的に）',
    '④ 💰 売上・客単価アップ施策（低コスト・低労力優先）',
    '   即実行可能な施策。回数券・物販・次回予約の改善ヒント',
    '⑤ 🛍 商材・技術情報',
    '   パーフェクトラッシュジャパン・大浴場など注目商材の最新情報',
    '⑥ 👥 採用・組織・フランチャイズ',
    '   採用成功事例・福利厚生アイデア・フランチャイズ動向（具体例付き）',
    '⑦ 📝 今日すぐ実行すべきアクション TOP3',
    '   優先順位付き。担当者レベルで実行できる具体的な粒度で',
    '',
    '重要: HTMLはインラインCSSのみ。見出しは色・サイズで視認性を高く。',
    'モバイルでも読みやすいレイアウト。各セクションは必ず過去レポートと異なる内容にすること。',
  ];

  var userPrompt = userLines.join('\n');

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:      'post',
    contentType: 'application/json',
    headers: {
      'x-api-key':         apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model:      'claude-opus-4-6',
      max_tokens: 7000,
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
      subject:       today + ' ' + focus.theme,
      slack_summary: '本日テーマ: ' + focus.theme + ' のレポートを送信しました',
      html:          '<p style="font-size:14px;line-height:1.8">' + rawText.replace(/\n/g, '<br>') + '</p>'
    };
  }

  var fullHtml = rpt_wrapEmail(parsed.html || '', parsed.subject || today + ' レポート', today, newsData.total, focus.theme);

  return {
    subject:      parsed.subject      || today + ' サロン日次レポート',
    slackSummary: parsed.slack_summary || '本日テーマ: ' + focus.theme,
    bodyHtml:     parsed.html          || '',
    html:         fullHtml
  };
}

// ================================================================
// メール本文HTML ラッパー
// ================================================================
function rpt_wrapEmail(bodyHtml, subject, today, newsCount, theme) {
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
    + today + '｜' + (theme || '') + '｜収集記事 ' + newsCount + '件｜AI分析レポート</p>\n'
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
// スプレッドシートに履歴保存（テーマ列追加）
// ================================================================
function rpt_saveHistory(report, theme) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('レポート履歴');

  if (!sh) {
    sh = ss.insertSheet('レポート履歴');
    // ヘッダー（5列目にテーマを追加）
    sh.appendRow(['送信日時', '件名', 'サマリー', '送信先', '本日のテーマ']);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(2, 260);
    sh.setColumnWidth(3, 300);
    sh.setColumnWidth(4, 160);
    sh.setColumnWidth(5, 180);
    sh.getRange('1:1').setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold');
  }

  sh.appendRow([
    new Date(),
    report.subject,
    report.slackSummary,
    RPT_TO,
    theme || ''
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

  // 毎朝8時（Asia/Tokyo）
  ScriptApp.newTrigger('generateDailyReport')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('トリガー設定完了: 毎朝8時に generateDailyReport を実行');

  var slack = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL') || '';
  try {
    rpt_postSlack(slack,
      'トリガー設定完了: 毎朝8時に ' + RPT_TO + ' へ自動配信します');
  } catch (e) {}
}
