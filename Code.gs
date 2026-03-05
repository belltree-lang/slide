/**************************************
 * AI研修スライド生成システム (Google Apps Script)
 * - Sidebar UI
 * - OpenAI Responses API
 * - Structured Output JSON (json_schema)
 * - Slides自動生成
 * - テンプレ対応
 * - Retry処理
 * - speakerNotes対応
 * - カリキュラム生成
 **************************************/

const CFG = {
  OPENAI_ENDPOINT: 'https://api.openai.com/v1/responses',
  DEFAULT_MODEL: 'gpt-4o-mini',
  APP_TITLE_PREFIX: '研修スライド：',
  DEFAULT_PAGE_COUNT: 15,
  MAX_PAGE_COUNT: 30,
  MIN_PAGE_COUNT: 5,
  REQUEST_TIMEOUT_MS: 60 * 1000,
  RETRY_MAX: 3,
  RETRY_BASE_SLEEP_MS: 800,
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('研修スライド生成')
    .addItem('サイドバーを開く', 'openSlideGenSidebar')
    .addSeparator()
    .addItem('テスト生成（固定テーマ/15枚）', 'testGenerateSlides_')
    .addToUi();
}

function openSlideGenSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('研修スライド生成');
  SpreadsheetApp.getUi().showSidebar(html);
}

function generateTrainingSlidesFromUi(input) {
  return generateTrainingSlidesFromUi_(input);
}

function testGenerateSlides_() {
  const result = generateTrainingSlidesFromUi_({
    title: 'ストレッチ研修（テスト）',
    theme: 'ストレッチの効果と安全なパートナーストレッチの現場運用',
    pageCount: 15,
    audience: '高齢者向けフィットネス施設のトレーナー',
    tone: '専門的だが現場で使える。安全性と標準化を重視。',
    templatePresentationId: '',
    curriculumLevel: '初級〜中級',
  });
  Logger.log(result);
}

function generateTrainingSlidesFromUi_(input) {
  const normalized = validateInput_(input);
  const apiKey = getApiKey_();
  const model = getModel_();

  const spec = buildSlideSpecWithAI_({
    apiKey,
    model,
    title: normalized.title,
    theme: normalized.theme,
    pageCount: normalized.pageCount,
    audience: normalized.audience,
    tone: normalized.tone,
    curriculumLevel: normalized.curriculumLevel,
  });

  const presentationId = createSlidesFromSpec_({
    spec,
    title: normalized.title,
    templatePresentationId: normalized.templatePresentationId,
  });

  return {
    ok: true,
    presentationId,
    presentationUrl: `https://docs.google.com/presentation/d/${presentationId}/edit`,
    slideCount: spec.slides.length,
    curriculum: spec.curriculum,
  };
}

function validateInput_(input) {
  if (!input) throw new Error('入力が空です');

  const title = String(input.title || '').trim();
  const theme = String(input.theme || '').trim();
  const pageCount = Number(input.pageCount || CFG.DEFAULT_PAGE_COUNT);
  const audience = String(input.audience || '').trim() || '研修参加トレーナー';
  const tone = String(input.tone || '').trim() || '実践重視';
  const templatePresentationId = String(input.templatePresentationId || '').trim();
  const curriculumLevel = String(input.curriculumLevel || '').trim() || '初級〜中級';

  if (!title) throw new Error('タイトルが未入力です');
  if (!theme) throw new Error('テーマが未入力です');

  if (!Number.isFinite(pageCount)) throw new Error('ページ数が不正です');
  if (pageCount < CFG.MIN_PAGE_COUNT || pageCount > CFG.MAX_PAGE_COUNT) {
    throw new Error(`ページ数は ${CFG.MIN_PAGE_COUNT}〜${CFG.MAX_PAGE_COUNT} の範囲にしてください`);
  }

  return {
    title,
    theme,
    pageCount,
    audience,
    tone,
    templatePresentationId,
    curriculumLevel,
  };
}

function buildSlideSpecWithAI_({ apiKey, model, title, theme, pageCount, audience, tone, curriculumLevel }) {
  const schema = getSlideSpecJsonSchema_(pageCount);
  const prompt = buildPrompt_({ title, theme, pageCount, audience, tone, curriculumLevel });

  const payload = {
    model,
    input: [{ role: 'user', content: [{ type: 'input_text', text: prompt }] }],
    text: {
      format: {
        type: 'json_schema',
        name: schema.name,
        schema: schema.schema,
        strict: true,
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
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: `Bearer ${apiKey}` },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
        timeout: CFG.REQUEST_TIMEOUT_MS,
      });

      const code = res.getResponseCode();
      const body = res.getContentText();

      if (code >= 200 && code < 300) return safeJsonParse_(body);

      const retryable = code === 429 || (code >= 500 && code <= 599);
      const msg = `OpenAI API error: HTTP ${code} body=${body}`;
      if (!retryable) throw new Error(msg);

      lastErr = new Error(msg);
    } catch (e) {
      lastErr = e;
    }

    if (attempt < CFG.RETRY_MAX) {
      const jitter = Math.floor(Math.random() * 250);
      const sleepMs = CFG.RETRY_BASE_SLEEP_MS * Math.pow(2, attempt) + jitter;
      Utilities.sleep(sleepMs);
    }
  }

  throw lastErr || new Error('OpenAI API呼び出しに失敗しました');
}

function extractJsonFromResponses_(json) {
  if (!json) throw new Error('OpenAI応答が空です');

  if (json.output_parsed && typeof json.output_parsed === 'object') {
    return json.output_parsed;
  }

  const texts = [];
  if (Array.isArray(json.output)) {
    json.output.forEach((outItem) => {
      if (!outItem || !Array.isArray(outItem.content)) return;
      outItem.content.forEach((part) => {
        if (part && typeof part.text === 'string') texts.push(part.text);
      });
    });
  }

  if (texts.length > 0) {
    const joined = texts.join('\n').trim();
    const clean = joined.replace(/```json|```/g, '').trim();
    return safeJsonParse_(clean);
  }

  if (typeof json.text === 'string') return safeJsonParse_(json.text);

  throw new Error('OpenAI応答からJSONを抽出できませんでした（レスポンス形式が想定外）');
}

function safeJsonParse_(s) {
  try {
    return JSON.parse(s);
  } catch (e) {
    throw new Error(`JSON parse failed: ${e.message}\nraw=${String(s).slice(0, 1000)}`);
  }
}

function getSlideSpecJsonSchema_(pageCount) {
  return {
    name: 'training_slides_schema',
    schema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        curriculum: {
          type: 'object',
          additionalProperties: false,
          properties: {
            courseTitle: { type: 'string', minLength: 1 },
            learningGoals: {
              type: 'array',
              minItems: 3,
              maxItems: 5,
              items: { type: 'string', minLength: 1 },
            },
            modules: {
              type: 'array',
              minItems: 3,
              maxItems: 8,
              items: {
                type: 'object',
                additionalProperties: false,
                properties: {
                  name: { type: 'string', minLength: 1 },
                  durationMinutes: { type: 'integer', minimum: 5, maximum: 180 },
                  objective: { type: 'string', minLength: 1 },
                },
                required: ['name', 'durationMinutes', 'objective'],
              },
            },
          },
          required: ['courseTitle', 'learningGoals', 'modules'],
        },
        slides: {
          type: 'array',
          minItems: pageCount,
          maxItems: pageCount,
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              title: { type: 'string', minLength: 1 },
              points: {
                type: 'array',
                minItems: 3,
                maxItems: 5,
                items: { type: 'string', minLength: 1 },
              },
              speakerNotes: { type: 'string', minLength: 1 },
            },
            required: ['title', 'points', 'speakerNotes'],
          },
        },
      },
      required: ['curriculum', 'slides'],
    },
  };
}

function validateSpec_(spec, pageCount) {
  if (!spec || !spec.curriculum || !Array.isArray(spec.slides)) {
    throw new Error('spec構造が不正です（curriculum/slides）');
  }

  if (spec.slides.length !== pageCount) {
    throw new Error(`スライド枚数が一致しません：期待=${pageCount} 実際=${spec.slides.length}`);
  }

  spec.slides.forEach((slide, idx) => {
    if (!slide.title || !Array.isArray(slide.points) || slide.points.length < 3 || !slide.speakerNotes) {
      throw new Error(`スライド${idx + 1}の構造が不正です`);
    }
  });
}

function buildPrompt_({ title, theme, pageCount, audience, tone, curriculumLevel }) {
  return `
あなたは研修設計と運動指導に精通したインストラクショナルデザイナーです。
以下の条件で「トレーナー研修」のカリキュラムとスライドを作成してください。

【タイトル】
${title}

【テーマ】
${theme}

【対象】
${audience}

【トーン】
${tone}

【難易度】
${curriculumLevel}

【必須要件】
- スライドは合計 ${pageCount} ページ
- 各スライドは points を 3〜5 個
- speakerNotes は登壇者がそのまま読める具体性で 2〜4 文
- curriculum.modules は時系列で、各moduleに所要時間(分)を設定
- 抽象論だけでなく、現場での注意点・禁忌・評価指標を含める
- 高齢者指導リスク(骨粗鬆症/循環器リスク/疼痛)への配慮を含める

【出力ルール】
- 必ず指定JSONスキーマに厳密準拠
- JSON以外の文字を出力しない
  `.trim();
}

function createSlidesFromSpec_({ spec, title, templatePresentationId }) {
  const finalTitle = `${CFG.APP_TITLE_PREFIX}${title}`;
  const presentation = templatePresentationId
    ? copyTemplatePresentation_(templatePresentationId, finalTitle)
    : SlidesApp.create(finalTitle);

  cleanupDefaultSlides_(presentation);

  spec.slides.forEach((slideSpec) => {
    const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    setTitleAndBody_(slide, slideSpec.title, slideSpec.points);

    const notesShape = slide.getNotesPage().getSpeakerNotesShape();
    notesShape.getText().setText(slideSpec.speakerNotes);
  });

  appendCurriculumSlide_(presentation, spec.curriculum);
  return presentation.getId();
}

function appendCurriculumSlide_(presentation, curriculum) {
  const slide = presentation.insertSlide(0, SlidesApp.PredefinedLayout.TITLE_AND_BODY);
  const moduleLines = curriculum.modules.map((m, i) => `${i + 1}. ${m.name}（${m.durationMinutes}分）: ${m.objective}`);
  const body = ['学習目標', ...curriculum.learningGoals.map((g) => `- ${g}`), '', 'モジュール構成', ...moduleLines].join('\n');
  setTitleAndBody_(slide, `カリキュラム概要: ${curriculum.courseTitle}`, body.split('\n'));

  const notesShape = slide.getNotesPage().getSpeakerNotesShape();
  notesShape.getText().setText('このスライドで本研修の全体像を説明し、受講者の期待値を揃えてください。');
}

function copyTemplatePresentation_(templateId, newName) {
  const file = DriveApp.getFileById(templateId);
  const copied = file.makeCopy(newName);
  return SlidesApp.openById(copied.getId());
}

function cleanupDefaultSlides_(presentation) {
  const slides = presentation.getSlides();
  if (slides.length !== 1) return;

  const first = slides[0];
  const shapes = first.getShapes();
  const hasText = shapes.some((shape) => {
    try {
      return shape.getText && shape.getText().asString().trim().length > 0;
    } catch (e) {
      return false;
    }
  });

  if (!hasText) first.remove();
}

function setTitleAndBody_(slide, title, points) {
  const titlePh = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
  const bodyPh = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
  const bodyText = points.map((point) => (point.startsWith('-') || point.match(/^\d+\./) ? point : `• ${point}`)).join('\n');

  if (titlePh) titlePh.asShape().getText().setText(title);

  if (bodyPh) {
    bodyPh.asShape().getText().setText(bodyText);
  } else {
    const box = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 60, 120, 600, 300);
    box.getText().setText(bodyText);
  }
}

function getApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) throw new Error('スクリプトプロパティ OPENAI_API_KEY が未設定です');
  return key;
}

function getModel_() {
  return PropertiesService.getScriptProperties().getProperty('OPENAI_MODEL') || CFG.DEFAULT_MODEL;
}
