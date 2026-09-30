/**
 * せんせいアシスト(仮想の児童生徒で支援を学ぶ) - GASバックエンド v1.1.2(公開版)
 * (v1.1:   チャートの形から架空の児童生徒のすがたを言葉にする describe_shape を追加。版の数字を自分用と合わせました)
 * (v1.1.2: Geminiが混雑(503)・回数制限(429)のとき、自動で待って再試行し、別のモデルに切り替える。
 *          スクリプト プロパティ GEMINI_MODELS に「モデル名,モデル名」と書けば、コードを直さずにモデルを差し替えられる。
 *          AIの返事の形が崩れていても受け取れるよう、返事を整えてからアプリに返す)
 *
 * 【大切な前提】
 * このアプリは、架空の児童生徒の特性を入れ、典型的に考えられる支援を想定するための道具です。
 * 実際の児童生徒のデータは入力しないでください。入力内容は Gemini API に送られます。
 *
 * 【セットアップ手順】
 * 1. https://script.google.com で新規プロジェクトを作成し、このファイルの中身を貼り付ける
 * 2. 「プロジェクトの設定」→「スクリプト プロパティ」で
 *    プロパティ名: GEMINI_API_KEY / 値: (Google AI Studioで取得したAPIキー)
 * 3. 「デプロイ」→「新しいデプロイ」→ウェブアプリ(実行:自分/アクセス:全員)でデプロイ
 *    ※コードを更新した場合は「デプロイを管理」→編集→新バージョンで再デプロイが必要
 * 4. Googleドキュメント出力機能を使うため、初回実行時に権限の承認を求められたら許可してください
 *    (ドキュメントは、デプロイした人自身のGoogleドライブに作成されます)
 * 5. 「アクセスできるユーザー:全員」にすると、URLを知っている人は誰でも呼び出せます。
 *    スクリプト プロパティに APP_KEY(合言葉)を設定し、アプリの「合言葉」欄に同じものを入れてください。
 *    合言葉が違う呼び出しは、AI にもドキュメント作成にも進まずに止まります。
 */

// 使えるモデル名は変わることがあります。うまく動かないときは、スクリプト プロパティ GEMINI_MODELS で差し替えられます。
const GEMINI_MODEL = 'gemini-3.5-flash';

// 権限確認用(初回だけ、これを実行して承認画面が出たら許可してください。
// 実行後、Googleドライブに「権限確認用テスト」というファイルができるので、あとで削除して構いません)
function authorizeTest() {
  const doc = DocumentApp.create('権限確認用テスト(削除してよい)');
  Logger.log(doc.getUrl());
}

function doGet(e) {
  return ContentService.createTextOutput(
    'せんせいアシスト バックエンド: 稼働中です。このURLはアプリからのPOSTアクセス専用なので、ブラウザで直接開いても機能は呼び出されません(これは正常な状態です)。'
  );
}

function doPost(e) {
  requestStartedAt = Date.now();
  try {
    if (!e || !e.postData) throw new Error('アプリからの呼び出しではありません(エディタで doPost を直接実行した場合はこのエラーになります)。');
    const data = JSON.parse(e.postData.contents);
    const appKey = PropertiesService.getScriptProperties().getProperty('APP_KEY');
    if (appKey && data.appKey !== appKey) {
      throw new Error('合言葉が違います。アプリの「合言葉」欄を確認してください。');
    }
    delete data.appKey;
    let result;
    switch (data.action) {
      case 'estimate_radar':
        result = { ok: true, values: normalizeValues(estimateRadar(data), (data.itemLabels || []).length) };
        break;
      case 'describe_shape':
        result = { ok: true, result: normalizeShapeResult(describeShape(data)) };
        break;
      case 'more_methods':
        result = { ok: true, methods: normalizeMethods(moreMethods(data)) };
        break;
      case 'comment_method':
        result = { ok: true, comment: commentMethod(data) };
        break;
      case 'export_doc':
        result = { ok: true, url: exportDoc(data) };
        break;
      default:
        const consultResult = consult(data);
        result = { ok: true, summary: consultResult.summary, detail: normalizeDetail(consultResult.detail) };
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: (err && err.message) ? err.message : String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ------------------------------------------------------------------
// AIの返事の形を整える(v1.1.2)
// 例: 配列のはずが {"values":[...]} で返ってくる、数が足りない、などを吸収する
// ------------------------------------------------------------------
function firstArray(x, keys) {
  if (Array.isArray(x)) return x;
  if (x && typeof x === 'object') {
    for (const k of keys) if (Array.isArray(x[k])) return x[k];
    const found = Object.keys(x).map(k => x[k]).find(v => Array.isArray(v));
    if (found) return found;
  }
  return [];
}

function normalizeValues(raw, n) {
  const arr = firstArray(raw, ['values', 'scores', 'items']).map(v => {
    const x = (v && typeof v === 'object') ? (v.value !== undefined ? v.value : v.score) : v;
    if (x === null || x === undefined || x === '') return 50;   // 空は「判断不可(50)」。0(支援が要る)にしない
    const num = Number(x);
    return isFinite(num) ? Math.max(0, Math.min(100, Math.round(num))) : 50;
  });
  if (!arr.length) throw new Error('AIの返事から数値を読み取れませんでした。もう一度押してください。');
  while (arr.length < n) arr.push(50);   // 足りない分は「判断不可(50)」で埋める
  return n ? arr.slice(0, n) : arr;
}

function normalizeMethods(raw) {
  const arr = firstArray(raw, ['methods', 'items', 'suggestions'])
    .map(m => (m && typeof m === 'object') ? (m.text || m.method || m.title || JSON.stringify(m)) : m)
    .map(m => String(m || '').trim())
    .filter(Boolean);
  if (!arr.length) throw new Error('AIの返事から手立てを読み取れませんでした。もう一度押してください。');
  return arr;
}

function normalizeShapeResult(raw) {
  const r = (raw && typeof raw === 'object' && !Array.isArray(raw)) ? raw : {};
  const guesses = firstArray(r, ['guesses', 'questions', 'items']).slice(0, 3).map(g => ({
    scene: String((g && g.scene) || '場面').slice(0, 12),
    text: String((g && (g.text || g.question)) || (typeof g === 'string' ? g : '')).trim()
  })).filter(g => g.text);
  const outline = String(r.outline || r.summary || '').trim();
  if (!outline && !guesses.length) throw new Error('AIの返事から文章を読み取れませんでした。もう一度押してください。');
  return { outline: outline, guesses: guesses };
}

function normalizeDetail(raw) {
  const r = (raw && typeof raw === 'object' && !Array.isArray(raw)) ? raw : {};
  const goals = firstArray(r, ['goals']).slice(0, 3).map(g => ({
    title: String((g && (g.title || g.goal)) || '').trim(),
    methods: firstArray(g, ['methods', 'steps']).map(m => String(m)).filter(Boolean),
    evaluation: String((g && g.evaluation) || '').trim()
  })).filter(g => g.title);
  return {
    tendency: String(r.tendency || '').trim(),
    goals: goals,
    accommodations: Array.isArray(r.accommodations) ? r.accommodations.join('\n') : String(r.accommodations || '').trim()
  };
}

// ------------------------------------------------------------------
// 共通:基本情報の文字列化
// ------------------------------------------------------------------
function basicInfoText(data) {
  return `※以下は実在の人物ではなく、架空の児童生徒の特性です。
呼び名: ${data.caseId || '(未入力)'}
学年: ${data.grade || ''}
在籍形態: ${data.placement || ''}
障害種別(選択): ${(data.disabilities || []).join('、') || '(未選択)'}
想定する診断名: ${data.diagnosisText || '(未記入)'}
指導する先生の得意な指導スタイル: ${(data.teacherStyles || []).join('、') || '(未選択)'}
指導する先生についての補足: ${data.teacherNote || '(未記入)'}
架空の児童生徒のすがた(言葉): ${data.observationText || '(未記入)'}`;
}

// チャート上の位置を言葉に置き換える(HTML側と同じ区切り)
function levelWord(v) {
  if (v < 20) return 'かなり支援が要る';
  if (v < 40) return '支援があればできる';
  if (v < 60) return '場面によってまちまち';
  if (v < 80) return 'だいたい自分でできる';
  return 'ほぼ自分でできる';
}

function categoryAverage(cat) {
  const items = cat.items || [];
  if (!items.length) return 50;
  return Math.round(items.reduce((a, it) => a + Number(it.value || 0), 0) / items.length);
}

function categoriesText(data) {
  return (data.categories || []).map(cat => {
    if (cat.touched === false) {
      return `【${cat.name}】未設定(想定に含まれていない区分。この区分については推測しないこと)`;
    }
    const itemsText = cat.items.map(it => {
      const chipsText = (it.selectedChips && it.selectedChips.length) ? `補足:${it.selectedChips.join('/')}` : '';
      const noteText = it.note ? `事実記録:${it.note}` : '';
      return `  - ${it.label}: ${it.value}/100 ${chipsText} ${noteText}`.trim();
    }).join('\n');
    return `【${cat.name}】\n${itemsText}`;
  }).join('\n\n');
}

const PRINCIPLES = `# 重要な原則(必ず守ること)
- 入力は実在の児童生徒ではなく、架空の児童生徒の特性である。目的は、このような特性をもつ児童生徒に典型的に考えられる支援を想定すること。特定の個人の見立てではなく、「このような特性の場合、典型的には〜のような支援が考えられる」という形で書く
- 入力に氏名・学校名など個人を特定しうる情報が含まれていても、出力でそれを繰り返さない
- 断定的な言い切り表現(「〜が原因である」「〜すべきである」等)は使わない。必ず「〜の可能性が考えられる」という仮説の形で書く
- 診断名や障害特性から性格や能力を決めつけない
- 数値はあくまで大まかな感触であり、精密な測定値ではないことを踏まえる
- 現場で実際に指導する先生が、無理なく実践でき、達成感を持てる内容を意識する
- 手立てを提案する際は、「指導する先生の得意な指導スタイル」が入力されていれば、それを活かせる手立てを優先的に盛り込む。ただし、先生の得意さを理由に本人の目標からズレた提案をしてはいけない
- 手立ては、抽象的な方針の言葉だけで終わらせない。「安心感を高める」のような一般論は現場の先生はすでに分かっている前提とする。必ず具体的な教材・道具の名前、声かけの言い回しの例、手順の流れなど、その日から真似できる具体性まで書くこと`;

// ------------------------------------------------------------------
// アクション1: 通常の相談(要点テキスト + 詳細JSON を別々に呼び出す)
// ------------------------------------------------------------------
function consult(data) {
  const basic = basicInfoText(data);
  const cats = categoriesText(data);

  const summaryPrompt = `あなたは、特別支援教育に関わる教員が、不特定の児童生徒を想定した仮想の児童生徒を題材に、支援の考え方を学ぶのを助けるアシスタントです。特定の児童生徒についての相談には使いません。

${PRINCIPLES}

# 基本情報
${basic}

# 自立活動6区分27項目の状態(0=困難が大きい, 100=できている・困難が少ない)
${cats}

# 依頼内容
研修や校内の話し合いで共有できる分量で、この特性に対して典型的に考えられる支援の要点を書いてください。条件:
- この特性で典型的に考えられることを、1〜2行で1つだけ
- 指導方針は、目標を箇条書きで3つ。それぞれ、目標の直後に「例:〜」という形で具体的な方法を一言添える
- 合計で15行以内におさめること。前置きの文章は書かない`;

  const summary = callGemini(summaryPrompt);

  const detailPrompt = `あなたは、特別支援教育に関わる教員が、不特定の児童生徒を想定した仮想の児童生徒を題材に、支援の考え方を学ぶのを助けるアシスタントです。特定の児童生徒についての相談には使いません。

${PRINCIPLES}

# 基本情報
${basic}

# 自立活動6区分27項目の状態(0=困難が大きい, 100=できている・困難が少ない)
${cats}

# 依頼内容
以下のJSON形式で、指導のアイデアを出力してください。
{
  "tendency": "この特性で典型的に考えられることの説明文(仮説として、根拠つきで2〜4文)",
  "goals": [
    { "title": "目標の文言", "methods": ["具体的な手立て1(教材名・声かけ例など具体的に)", "具体的な手立て2"], "evaluation": "評価の視点" }
  ],
  "accommodations": "合理的配慮・支援機器(ICT・AT)の候補の説明文"
}
goalsは優先度の高い順に最大3つ。`;

  const detail = callGeminiJson(detailPrompt);
  return { summary: summary, detail: detail };
}

// ------------------------------------------------------------------
// アクション2: 特定の目標に対する追加・差し替えの手立て提案
// ------------------------------------------------------------------
function moreMethods(data) {
  const count = data.count || 2;
  const prompt = `あなたは、特別支援教育に関わる教員が、不特定の児童生徒を想定した仮想の児童生徒を題材に、支援の考え方を学ぶのを助けるアシスタントです。特定の児童生徒についての相談には使いません。

${PRINCIPLES}

# 基本情報
${basicInfoText(data)}

# 対象の目標
${data.goalTitle || ''}

# すでに提案済み・採用中の手立て(これらとは違う具体的な方法を提案すること)
${(data.existingMethods || []).map((m, i) => `${i + 1}. ${m}`).join('\n')}

# 依頼内容
上記の目標に対する、追加の具体的な手立てを${count}個提案してください。すでに提案済みのものと重複しないこと。

# 出力形式(厳守)
説明・前置きは一切書かず、手立ての文字列だけを含むJSON配列のみを出力すること。
例: ["手立てA(具体的に)", "手立てB(具体的に)"]`;

  return callGeminiJson(prompt);
}

// ------------------------------------------------------------------
// アクション3: 先生が思いついた案へのコメント
// ------------------------------------------------------------------
function commentMethod(data) {
  const prompt = `あなたは、特別支援教育に関わる教員が、不特定の児童生徒を想定した仮想の児童生徒を題材に、支援の考え方を学ぶのを助けるアシスタントです。特定の児童生徒についての相談には使いません。

${PRINCIPLES}

# 基本情報
${basicInfoText(data)}

# 対象の目標
${data.goalTitle || ''}

# 先生が考えている案
${data.proposedMethod || ''}

# 依頼内容
この案について、コーディネーターとしてのコメントを書いてください。良い点、気をつけたい点、必要であれば少し調整した案を、断定せず仮説の形で、3〜5行程度で簡潔に書いてください。前置きは不要です。`;

  return callGemini(prompt);
}

// ------------------------------------------------------------------
// アクション4: 架空の児童生徒の様子(言葉)から27項目を推定
// ------------------------------------------------------------------
function estimateRadar(data) {
  const labels = data.itemLabels || [];
  const listText = labels.map((l, i) => `${i + 1}. ${l}`).join('\n');
  const prompt = `以下は、典型的な支援を想定するためにつくった、架空の児童生徒の様子を書いた文章です。実在の人物ではありません。

# 架空の児童生徒の様子
${data.text || '(記入なし)'}

# 依頼内容
上記の内容から、以下の${labels.length}項目それぞれについて、0〜100の数値で状態を推定してください。
0=困難が大きい、100=できている・困難が少ない、を意味します。
言葉の中に手がかりが全くなく判断できない項目は50(中間・判断不可)としてください。

# 対象項目(この順番のまま)
${listText}

# 出力形式(厳守)
${labels.length}個の数値だけを含むJSON配列のみを出力すること。`;

  return callGeminiJson(prompt);
}

// ------------------------------------------------------------------
// アクション4b: チャートの形から児童生徒のすがたを言葉にする(v1.1)
// ------------------------------------------------------------------
function describeShape(data) {
  const shapeText = (data.shape || []).map(s => s.touched
    ? `- ${s.name}: ${s.word}(目安 ${s.value}/100)`
    : `- ${s.name}: 未設定(この区分は想定していない)`
  ).join('\n');
  const notesText = (data.itemNotes || []).length
    ? (data.itemNotes || []).map(n => `- ${n}`).join('\n')
    : '(なし)';

  const prompt = `あなたは、特別支援教育に関わる教員が、不特定の児童生徒を想定した仮想の児童生徒を題材に、支援の考え方を学ぶのを助けるアシスタントです。特定の児童生徒についての相談には使いません。
教員が、典型的な支援を考えるために想定した架空の児童生徒の全体像を、自立活動6区分のレーダーチャートの「形」として指で描きました。実在の人物ではありません。
中心に近いほど「かなり支援が要る」、外側ほど「ほぼ自分でできる」を表します。
これは測定値ではなく、想定のおおまかなイメージです。

# 描かれた形
${shapeText}

# 項目ごとの補足(あれば)
${notesText}

# 基本情報
学年: ${data.grade || ''}
在籍形態: ${data.placement || ''}
障害種別(選択): ${(data.disabilities || []).join('、') || '(未選択)'}
具体的な診断名: ${data.diagnosisText || '(未記入)'}
すでに書かれている言葉: ${data.observationText || '(なし)'}

# 依頼内容
この形から想像される架空の児童生徒のすがたを、使う人に「こんな感じですか?」と確かめてもらうための短い文章にしてください。
- outline: 2〜3文。外側に大きく描かれた区分(できていること・強み)から先に書き、次に内側の区分に触れる。「〜のような児童生徒が考えられます」と仮説の形で書く
- guesses: ちょうど3つ。それぞれ学校の具体的な場面を1つ選び(例: 朝の会、移動教室、休み時間、授業中の課題、給食、行事の練習)、その場面で見られそうな様子を、観察できる行動の言葉で1文にする。「不安が強い」のような抽象語ではなく「予定が変わると手が止まり、先生の顔を見る」のように書く。文末は「〜ではないですか?」「〜ということはありませんか?」の問いかけにする
- 未設定の区分については書かない・推測しない
- 障害種別や診断名は場面を具体的にするためだけに使い、そこから性格や能力を決めつけない
- すでに書かれている言葉があれば、それと矛盾しないようにする

# 出力形式(厳守)
{"outline": "文章", "guesses": [{"scene": "場面名(8文字以内)", "text": "問いかけの1文"}]}`;

  return callGeminiJson(prompt);
}

// ------------------------------------------------------------------
// アクション5: Googleドキュメントとして出力
// ------------------------------------------------------------------
function exportDoc(data) {
  const detail = data.detail || {};
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const title = `せんせいアシスト_典型的な支援の想定(架空)_${data.caseId || 'ケース'}_${today}`;
  const doc = DocumentApp.create(title);
  const body = doc.getBody();

  body.appendParagraph('せんせいアシスト 典型的な支援の想定(架空の児童生徒)').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('※架空の児童生徒の特性から、典型的に考えられる支援を想定したものです。実在の児童生徒の見立てではありません。');
  body.appendParagraph(`呼び名: ${data.caseId || ''} / 学年: ${data.grade || ''} / 作成日: ${today}`);
  body.appendParagraph(`障害種別: ${(data.disabilities || []).join('、')}`);
  body.appendParagraph('');

  const shapeLines = (data.categories || [])
    .filter(cat => cat.touched !== false)
    .map(cat => `${cat.name}: ${levelWord(categoryAverage(cat))}`);
  if (data.observationText || shapeLines.length) {
    body.appendParagraph('架空の児童生徒のすがた(想定)').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    shapeLines.forEach(l => body.appendListItem(l).setGlyphType(DocumentApp.GlyphType.BULLET));
    if (data.observationText) body.appendParagraph(data.observationText);
    body.appendParagraph('※測定値ではなく、想定のおおまかなイメージです。');
  }

  body.appendParagraph('想定される典型的な支援の要点').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(data.summary || '');

  if (detail.tendency) {
    body.appendParagraph('この特性で典型的に考えられること').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(detail.tendency);
  }

  if (detail.goals && detail.goals.length) {
    body.appendParagraph('指導方針').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    detail.goals.forEach((g, i) => {
      body.appendParagraph(`${i + 1}. 目標: ${g.title}`).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      (g.methods || []).forEach(m => body.appendListItem(String(m)).setGlyphType(DocumentApp.GlyphType.BULLET));
      if (g.evaluation) body.appendParagraph(`評価の視点: ${g.evaluation}`);
    });
  }

  if (detail.accommodations) {
    body.appendParagraph('合理的配慮・支援機器の候補').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(detail.accommodations);
  }

  doc.saveAndClose();
  return doc.getUrl();
}

// ------------------------------------------------------------------
// Gemini呼び出し(v1.1.2: 混雑時の自動再試行とモデル切り替え)
// ------------------------------------------------------------------
// 上から順に試す。1つ目が混雑(503)やレート制限(429)なら次のモデルへ切り替える。
// スクリプトプロパティ GEMINI_MODELS に「モデル名,モデル名」と書けば、コードを直さずに差し替えられる。
const DEFAULT_MODELS = [GEMINI_MODEL, 'gemini-flash-latest', 'gemini-flash-lite-latest'];
const RETRY_WAITS_MS = [3000];         // 同じモデルでの再試行は1回だけ(混雑が続くなら早めに別モデルへ切り替える)
const TIME_BUDGET_MS = 270000;         // 1回のボタン操作で新しい呼び出しを始めてよい時間(GASの上限6分より手前)
const ONE_CALL_MAX_MS = 65000;         // Geminiへの1回の呼び出しにかかりうる最長時間の見込み
let requestStartedAt = Date.now();

function modelList() {
  const prop = PropertiesService.getScriptProperties().getProperty('GEMINI_MODELS');
  const list = prop ? prop.split(',').map(s => s.trim()).filter(Boolean) : DEFAULT_MODELS;
  return list.filter((m, i) => list.indexOf(m) === i);
}

function timeLeft() {
  return TIME_BUDGET_MS - (Date.now() - requestStartedAt);
}

function callGemini(prompt) {
  return callGeminiRaw(prompt, false);
}

// JSONで返してもらい、読み取ったオブジェクトを返す。形が崩れていたら1回だけ頼み直す
function callGeminiJson(prompt) {
  let lastErr;
  for (let i = 0; i < 2; i++) {
    const text = callGeminiRaw(prompt, true);
    try {
      return parseJsonLoose(text);
    } catch (e) {
      lastErr = e;
    }
  }
  throw new Error('AIの返事をデータとして読み取れませんでした。もう一度押してください。(' + String(lastErr).slice(0, 80) + ')');
}

function parseJsonLoose(text) {
  const t = String(text).replace(/```json|```/g, '').trim();
  try {
    return JSON.parse(t);
  } catch (e) {
    const a = t.search(/[\[{]/);
    const b = Math.max(t.lastIndexOf(']'), t.lastIndexOf('}'));
    if (a >= 0 && b > a) return JSON.parse(t.slice(a, b + 1));
    throw e;
  }
}

function callGeminiRaw(prompt, forceJson) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEYが設定されていません。スクリプトプロパティを確認してください。');

  const payload = { contents: [{ role: 'user', parts: [{ text: prompt }] }] };
  if (forceJson) payload.generationConfig = { responseMimeType: 'application/json' };

  const notes = [];
  const models = modelList();
  for (let mi = 0; mi < models.length; mi++) {
    const model = models[mi];
    for (let attempt = 0; attempt <= RETRY_WAITS_MS.length; attempt++) {
      if (timeLeft() < ONE_CALL_MAX_MS) {
        throw new Error('Geminiが混み合っていて、時間内に返事がありませんでした。1〜2分おいてからもう一度押してください。(' + notes.join(' / ') + ')');
      }
      const r = fetchGemini(model, payload, apiKey);

      if (r.code === 200) {
        return extractText(r.body, model);
      }
      notes.push(`${model}:${r.code}`);

      if (r.code === 400) {
        throw new Error('Geminiへの依頼の形が正しくありません(400)。' + errorMessage(r));
      }
      if (r.code === 401 || r.code === 403) {
        throw new Error('APIキーが使えない状態です(' + r.code + ')。Google AI Studioでキーを確認してください。' + errorMessage(r));
      }
      if (r.code === 404) break;            // このモデル名が無い → 次のモデルへ
      if (r.code === 429) break;            // このモデルの無料枠を使い切った → 次のモデルへ
      // 500 / 502 / 503 / 504 など一時的な障害 → 少し待って同じモデルで再試行
      if (attempt < RETRY_WAITS_MS.length) {
        Utilities.sleep(Math.min(RETRY_WAITS_MS[attempt] + Math.floor(Math.random() * 1000), Math.max(0, timeLeft() - ONE_CALL_MAX_MS)));
      }
    }
  }
  if (notes.length && notes.every(n => /:404$/.test(n))) {
    throw new Error('Geminiのモデル名が見つかりませんでした(廃止された可能性があります)。スクリプトプロパティ GEMINI_MODELS に、今使えるモデル名を書いてください。(' + notes.join(' / ') + ')');
  }
  if (notes.length && notes.every(n => /:429$/.test(n))) {
    throw new Error('Geminiの無料枠の回数制限に達しました。1〜2分おいてからもう一度押してください。(' + notes.join(' / ') + ')');
  }
  throw new Error('Geminiが混み合っていて、切り替え先のモデルも含めて返事がありませんでした。1〜2分おいてからもう一度押してください。(' + notes.join(' / ') + ')');
}

function fetchGemini(model, payload, apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;
  let res;
  try {
    res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    return { code: 503, body: null, raw: String(e) };   // 通信そのものの失敗も一時的な障害として扱う
  }
  const raw = res.getContentText();
  let body = null;
  try { body = JSON.parse(raw); } catch (e) { /* エラーページ(HTML)などJSONでない返事 */ }
  let code = res.getResponseCode();
  if (code === 200 && !body) code = 502;   // 200なのに中身が読めない → 一時的な障害扱い
  return { code: code, body: body, raw: raw };
}

function errorMessage(r) {
  const m = r.body && r.body.error && r.body.error.message;
  return m ? ' ' + String(m).slice(0, 160) : '';
}

function extractText(body, model) {
  const cand = body.candidates && body.candidates[0];
  if (!cand) {
    const reason = body.promptFeedback && body.promptFeedback.blockReason;
    throw new Error('AIが返事をしませんでした' + (reason ? `(理由: ${reason})` : '') + '。入力の言葉を少し変えて試してください。');
  }
  const parts = (cand.content && cand.content.parts) || [];
  const text = parts.filter(p => !p.thought).map(p => p.text || '').join('').trim();
  if (!text) {
    const reason = cand.finishReason || '不明';
    if (reason === 'MAX_TOKENS') throw new Error('AIの返事が長すぎて途中で切れました。もう一度押してください。');
    if (reason === 'SAFETY' || reason === 'PROHIBITED_CONTENT' || reason === 'BLOCKLIST') {
      throw new Error('AIの安全フィルターで返事が止められました。入力の言葉を少し変えて試してください。');
    }
    throw new Error(`AIの返事が空でした(${reason})。もう一度押してください。`);
  }
  return text;
}
