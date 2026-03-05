/**************************************
 * Strong Slide Generator (GAS)
 * - OpenAI Responses API + Structured Outputs (json_schema)
 * - Sidebar UI
 * - Template apply (optional)
 * - Retries + robust errors
 **************************************/

const CFG = {
  OPENAI_ENDPOINT: "https://api.openai.com/v1/responses",
  DEFAULT_MODEL: "gpt-4o-mini",
  APP_TITLE_PREFIX: "研修スライド：",
  DEFAULT_PAGE_COUNT: 15,
  MAX_PAGE_COUNT: 30,      // コスト暴騰防止
  MIN_PAGE_COUNT: 5,
  REQUEST_TIMEOUT_MS: 60 * 1000,
  RETRY_MAX: 3,
  RETRY_BASE_SLEEP_MS: 800,
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("研修スライド生成")
    .addItem("サイドバーを開く", "openSlideGenSidebar")
    .addSeparator()
    .addItem("テスト生成（固定テーマ/15枚）", "testGenerateSlides_")
    .addToUi();
}

function openSlideGenSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("研修スライド生成");
  SpreadsheetApp.getUi().showSidebar(html);
}

function testGenerateSlides_() {
  const input = {
    title: "ストレッチ研修（テスト）",
    theme: "ストレッチの効果について／トレーナーとしてパートナーストレッチを行う上での専門的ノウハウと考え方の共有",
    pageCount: 15,
    audience: "高齢者向けフィットネス施設のトレーナー（鍼灸師・健康運動指導士含む）",
    tone: "専門的だが現場で使える。安全性と標準化を重視。",
    templatePresentationId: "",
  };
  const result = generateTrainingSlidesFromUi_(input);
  Logger.log(result);
}

/**
 * Sidebar -> GAS 呼び出し口
 */
function generateTrainingSlidesFromUi_(input) {
  validateInput_(input);

  const apiKey = getApiKey_();
  const model = getModel_();

  // 1) AIへ「厳密なJSON schema」で出力させる
  const spec = buildSlideSpecWithAI_({
    apiKey,
    model,
    title: input.title,
    theme: input.theme,
    pageCount: input.pageCount,
    audience: input.audience,
    tone: input.tone,
  });

  // 2) Slides生成（テンプレIDがあれば複製して統一デザイン）
  const presentationId = createSlidesFromSpec_({
    spec,
    title: input.title,
    templatePresentationId: input.templatePresentationId,
  });

  return {
    ok: true,
    presentationId,
    presentationUrl: `https://docs.google.com/presentation/d/${presentationId}/edit`,
    slideCount: spec.slides.length,
  };
}

/**************************************
 * Input validation
 **************************************/
function validateInput_(input) {
  if (!input) throw new Error("入力が空です");

  const title = String(input.title || "").trim();
  const theme = String(input.theme || "").trim();
  const pageCount = Number(input.pageCount || CFG.DEFAULT_PAGE_COUNT);

  if (!title) throw new Error("タイトルが未入力です");
  if (!theme) throw new Error("テーマが未入力です");

  if (!Number.isFinite(pageCount)) throw new Error("ページ数が不正です");
  if (pageCount < CFG.MIN_PAGE_COUNT || pageCount > CFG.MAX_PAGE_COUNT) {
    throw new Error(`ページ数は ${CFG.MIN_PAGE_COUNT}〜${CFG.MAX_PAGE_COUNT} の範囲にしてください`);
  }
}

/**************************************
 * OpenAI call (Responses API)
 **************************************/
function buildSlideSpecWithAI_({ apiKey, model, title, theme, pageCount, audience, tone }) {
  const schema = getSlideSpecJsonSchema_(pageCount); // ★必ず additionalProperties:false を含む

  const prompt = buildPrompt_(title, theme, pageCount, audience, tone);

  // Responses API Structured Outputs（現行仕様：text.format）
  const payload = {
    model,
    input: [
      {
        role: "user",
        content: [
          { type: "input_text", text: prompt }
        ],
      },
    ],
    text: {
      format: {
        type: "json_schema",
        name: schema.name,     // ★必須
        schema: schema.schema, // ★ここが重要：additionalProperties:false を含むスキーマをそのまま渡す
        strict: true,          // ★可能なら strict をON（効く）
      },
    },
    temperature: 0.4,
  };

  const json = fetchOpenAIWithRetry_({
    url: CFG.OPENAI_ENDPOINT,
    apiKey,
    payload,
  });

  const parsed = extractJsonFromResponses_(json);
  validateSpec_(parsed, pageCount);
  return parsed;
}

function fetchOpenAIWithRetry_({ url, apiKey, payload }) {
  let lastErr = null;

  for (let attempt = 0; attempt <= CFG.RETRY_MAX; attempt++) {
    try {
      const res = UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        headers: { Authorization: `Bearer ${apiKey}` },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
        timeout: CFG.REQUEST_TIMEOUT_MS,
      });

      const code = res.getResponseCode();
      const body = res.getContentText();

      if (code >= 200 && code < 300) return safeJsonParse_(body);

      const retryable = (code === 429) || (code >= 500 && code <= 599);
      const msg = `OpenAI API error: HTTP ${code} body=${body}`;
      if (!retryable) throw new Error(msg);

      lastErr = new Error(msg);
      Utilities.sleep(CFG.RETRY_BASE_SLEEP_MS * Math.pow(2, attempt));
    } catch (e) {
      lastErr = e;
      Utilities.sleep(CFG.RETRY_BASE_SLEEP_MS * Math.pow(2, attempt));
    }
  }

  throw lastErr || new Error("OpenAI API呼び出しに失敗しました");
}

/**
 * Responses APIの返却からJSONを抽出（形式揺れに強い版）
 */
function extractJsonFromResponses_(json) {
  if (!json) throw new Error("OpenAI応答が空です");

  // 1) すでにパース済みっぽいキーが来るケース（環境差）
  if (json.output_parsed && typeof json.output_parsed === "object") return json.output_parsed;

  // 2) output 配列から output_text を拾う
  const texts = [];
  if (Array.isArray(json.output)) {
    json.output.forEach((o) => {
      if (!o || !Array.isArray(o.content)) return;
      o.content.forEach((c) => {
        // Responses API: type が output_text / summary_text 等
        if (c && typeof c.text === "string") texts.push(c.text);
      });
    });
  }

  if (texts.length > 0) {
    const joined = texts.join("\n").trim();
    const clean = joined.replace(/```json|```/g, "").trim();
    return safeJsonParse_(clean);
  }

  // 3) 予備：top-level に text が来るケース
  if (typeof json.text === "string") return safeJsonParse_(json.text);

  throw new Error("OpenAI応答からJSONを抽出できませんでした（レスポンス形式が想定外）");
}

function safeJsonParse_(s) {
  try {
    return JSON.parse(s);
  } catch (e) {
    throw new Error(`JSON parse failed: ${e.message}\nraw=${String(s).slice(0, 800)}`);
  }
}

/**************************************
 * Spec schema / validation
 **************************************/
function getSlideSpecJsonSchema_(pageCount) {
  return {
    name: "slides_schema",
    schema: {
      type: "object",
      additionalProperties: false,
      properties: {
        slides: {
          type: "array",
          minItems: pageCount,
          maxItems: pageCount,
          items: {
            type: "object",
            additionalProperties: false,
            properties: {
              title: { type: "string", minLength: 1 },
              points: {
                type: "array",
                minItems: 3,
                maxItems: 5,
                items: {
                  type: "string",
                  minLength: 1
                }
              }
            },
            required: ["title", "points"]
          }
        }
      },
      required: ["slides"]
    }
  };
}

function validateSpec_(spec, pageCount) {
  if (!spec || !Array.isArray(spec.slides)) throw new Error("spec.slides が不正です");
  if (spec.slides.length !== pageCount) {
    throw new Error(`スライド枚数が一致しません：期待=${pageCount} 実際=${spec.slides.length}`);
  }
  spec.slides.forEach((s, idx) => {
    if (!s.title || !Array.isArray(s.points) || s.points.length < 3) {
      throw new Error(`スライド${idx + 1}の構造が不正です`);
    }
  });
}

/**************************************
 * Prompt (品質の8割はここ)
 **************************************/
function buildPrompt_(title, theme, pageCount, audience, tone) {
  return `
あなたはフィットネス指導・運動生理学・リハビリテーション・高齢者運動指導に精通した専門家です。
以下の条件でトレーナー研修用スライドを作成してください。

【タイトル】
${title}

【テーマ】
${theme}

【対象】
${audience}

【トーン】
${tone}

【目的】
- パートナーストレッチを「安全かつ再現性高く」提供するための共通理解を作る
- 効果（何が変わるか）と限界（何は変わらないか）を区別する
- 禁忌・注意点・現場の失敗パターンを明確化する
- 評価→実施→再評価の流れで標準化する

【必須要件】
- ${pageCount}ページ
- 各ページ：箇条書き 3〜5点（1点は短文）
- 抽象論のみは禁止（現場で指示に落ちる粒度）
- 高齢者前提：骨粗鬆症、人工関節、循環器リスク、疼痛、聴力低下も想定
- 生理学：筋紡錘、ゴルジ腱器官、伸張反射、ストレッチ耐性、温熱/反復の効果
- 手技ノウハウ：ポジショニング、固定、テコ、力の方向、コミュニケーション、同意形成
- 現場運用：短時間で実施、声かけ、評価指標（可動域/痛み/歩行/立ち上がり）

【構成ガイド（自由だが必ず網羅）】
1) 導入：目的/ゴール
2) ストレッチの定義と誤解（柔軟性=正義、は危険）
3) 効果（短期・中期・長期）
4) 生理学（反射と安全）
5) セルフ vs パートナーの違い
6) 安全原則と禁忌
7) 実施手順（評価→実施→再評価）
8) よくある失敗と修正
9) 現場運用（時間/声かけ/記録）
10) ケース（股関節/ハム/肩甲帯など）とまとめ

【出力】
指定されたJSONスキーマに厳密に従い、JSONのみを出力してください。
`.trim();
}

/**************************************
 * Slides generation
 **************************************/
function createSlidesFromSpec_({ spec, title, templatePresentationId }) {
  const finalTitle = `${CFG.APP_TITLE_PREFIX}${title}`;

  const presentation = templatePresentationId
    ? copyTemplatePresentation_(templatePresentationId, finalTitle)
    : SlidesApp.create(finalTitle);

  cleanupDefaultSlides_(presentation);

  spec.slides.forEach((s) => {
    const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);

    setTitleAndBody_(slide, s.title, s.points);

    if (s.speakerNotes) {
      slide.getNotesPage().getSpeakerNotesShape().getText().setText(s.speakerNotes);
    } else {
      slide.getNotesPage().getSpeakerNotesShape().getText().setText(
        `話すポイント：${s.points.join(" / ")}`
      );
    }

    if (s.imagePrompt) {
      const current = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
      slide.getNotesPage().getSpeakerNotesShape().getText().setText(
        `${current}\n\n[図案] ${s.imagePrompt}`
      );
    }
  });

  return presentation.getId();
}

function copyTemplatePresentation_(templateId, newName) {
  const file = DriveApp.getFileById(templateId);
  const copied = file.makeCopy(newName);
  return SlidesApp.openById(copied.getId());
}

function cleanupDefaultSlides_(presentation) {
  const slides = presentation.getSlides();
  if (slides.length === 1) {
    const s = slides[0];
    const shapes = s.getShapes();
    const hasText = shapes.some(sh => {
      try { return sh.getText && sh.getText().asString().trim().length > 0; } catch (e) { return false; }
    });
    if (!hasText) s.remove();
  }
}

function setTitleAndBody_(slide, title, points) {
  const titlePh = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
  const bodyPh = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);

  if (titlePh) titlePh.asShape().getText().setText(title);

  if (bodyPh) {
    bodyPh.asShape().getText().setText(points.map(p => `• ${p}`).join("\n"));
  } else {
    const box = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 60, 120, 600, 300);
    box.getText().setText(points.map(p => `• ${p}`).join("\n"));
  }
}

/**************************************
 * Secrets / config
 **************************************/
function getApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!key) throw new Error("スクリプトプロパティ OPENAI_API_KEY が未設定です");
  return key;
}

function getModel_() {
  return (
    PropertiesService.getScriptProperties().getProperty("OPENAI_MODEL") ||
    CFG.DEFAULT_MODEL
  );
}

function generateTrainingSlidesFromUi(input) {
  // Sidebar から呼ぶ公開関数（_なし）
  return generateTrainingSlidesFromUi_(input);
}
