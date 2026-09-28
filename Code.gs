/**
 * せんせいアシスト(仮想の児童生徒で支援を学ぶ) - GASバックエンド v3.4(公開版)
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

// 使えるモデル名は変わることがあります。エラーになったら 'gemini-flash-latest' などに書きかえてください。
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
  try {
    const data = JSON.parse(e.postData.contents);
    const appKey = PropertiesService.getScriptProperties().getProperty('APP_KEY');
    if (appKey && data.appKey !== appKey) {
      throw new Error('合言葉が違います。アプリの「合言葉」欄を確認してください。');
    }
    delete data.appKey;
    let result;
    switch (data.action) {
      case 'estimate_radar':
        result = { ok: true, values: estimateRadar(data) };
        break;
      case 'more_methods':
        result = { ok: true, methods: moreMethods(data) };
        break;
      case 'comment_method':
        result = { ok: true, comment: commentMethod(data) };
        break;
      case 'export_doc':
        result = { ok: true, url: exportDoc(data) };
        break;
      default:
        const consultResult = consult(data);
        result = { ok: true, summary: consultResult.summary, detail: consultResult.detail };
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
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
指導する先生についての補足: ${data.teacherNote || '(未記入)'}`;
}

function categoriesText(data) {
  return (data.categories || []).map(cat => {
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
  "tendency": "この特性で典型的に考えられることの説明文(根拠つきで2〜4文)",
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
// アクション4: 見取りの言葉から27項目を推定
// ------------------------------------------------------------------
function estimateRadar(data) {
  const labels = data.itemLabels || [];
  const listText = labels.map((l, i) => `${i + 1}. ${l}`).join('\n');
  const prompt = `以下は、典型的な支援を想定するためにつくった、架空の児童生徒の様子を書いた文章です。実在の人物ではありません。

# 観察・見取りの言葉
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
// アクション5: Googleドキュメントとして出力
// ------------------------------------------------------------------
function exportDoc(data) {
  const detail = data.detail || {};
  const title = `せんせいアシスト_典型的な支援の想定(架空)_${data.caseId || 'ケース'}_${data.visitDate || ''}`;
  const doc = DocumentApp.create(title);
  const body = doc.getBody();

  body.appendParagraph('せんせいアシスト 典型的な支援の想定(架空の児童生徒)').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('※架空の児童生徒の特性から、典型的に考えられる支援を想定したものです。実在の児童生徒の見立てではありません。');
  body.appendParagraph(`呼び名: ${data.caseId || ''} / 学年: ${data.grade || ''} / 作成日: ${data.visitDate || ''}`);
  body.appendParagraph(`障害種別: ${(data.disabilities || []).join('、')}`);
  body.appendParagraph('');

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
      (g.methods || []).forEach(m => body.appendListItem(m).setGlyphType(DocumentApp.GlyphType.BULLET));
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
// Gemini呼び出し(通常テキスト)
// ------------------------------------------------------------------
function callGemini(prompt) {
  return callGeminiRaw(prompt, false);
}

// Gemini呼び出し(JSON強制・パース済みで返す)
function callGeminiJson(prompt) {
  const text = callGeminiRaw(prompt, true);
  return JSON.parse(text);
}

function callGeminiRaw(prompt, forceJson) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEYが設定されていません。スクリプトプロパティを確認してください。');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;
  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  if (forceJson) {
    payload.generationConfig = { responseMimeType: 'application/json' };
  }
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const body = JSON.parse(response.getContentText());

  if (code === 429) {
    throw new Error('APIのレート制限に達しました(無料枠の場合、しばらく待ってから再試行してください)。');
  }
  if (code !== 200) {
    throw new Error('Gemini APIエラー: ' + JSON.stringify(body));
  }

  const text = body.candidates && body.candidates[0] && body.candidates[0].content
    ? body.candidates[0].content.parts.map(p => p.text).join('')
    : '(応答を取得できませんでした)';
  return text;
}
