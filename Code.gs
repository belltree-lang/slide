/**************************************
 * Professional Trainer Slide Generator (Google Apps Script)
 * Multi-stage AI pipeline:
 * Theme -> Curriculum -> Slide Structure -> Slide Content -> Google Slides
 **************************************/

const CFG = {
  OPENAI_ENDPOINT: 'https://api.openai.com/v1/responses',
  DEFAULT_MODEL: 'gpt-4o-mini',
  APP_TITLE_PREFIX: 'Training Slides: ',
  DEFAULT_PAGE_COUNT: 12,
  MIN_PAGE_COUNT: 10,
  MAX_PAGE_COUNT: 20,
  REQUEST_TIMEOUT_MS: 60 * 1000,
  RETRY_MAX: 3,
  RETRY_BASE_SLEEP_MS: 900,
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Training Slide Generator')
    .addItem('Open Sidebar', 'openSlideGenSidebar')
    .addToUi();
}

function openSlideGenSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Training Slide Generator');
  SpreadsheetApp.getUi().showSidebar(html);
}

function generateTrainingSlidesFromUi(input) {
  const normalized = validateInput_(input);
  const apiKey = getApiKey_();
  const model = getModel_();

  const curriculum = generateCurriculum({
    apiKey,
    model,
    theme: normalized.theme,
    audience: normalized.audience,
    pageCount: normalized.pageCount,
  });

  const structure = generateSlideStructure({
    apiKey,
    model,
    sections: curriculum.sections,
    pageCount: normalized.pageCount,
  });

  const slideBlueprints = buildSlideBlueprints_(structure.slides, normalized.pageCount);
  const slideContents = slideBlueprints.map((blueprint, index) => {
    const content = generateSlideContent({
      apiKey,
      model,
      section: blueprint.section,
      slideIndex: index + 1,
      slideCount: normalized.pageCount,
      audience: normalized.audience,
      tone: normalized.tone,
    });

    return {
      section: blueprint.section,
      title: content.title,
      points: content.points,
      trainer_tips: content.trainer_tips,
      common_mistakes: content.common_mistakes,
      speaker_notes: content.speaker_notes,
    };
  });

  const presentationId = createSlides({
    title: normalized.title,
    templatePresentationId: normalized.templatePresentationId,
    curriculum,
    slides: slideContents,
  });

  return {
    ok: true,
    presentationId,
    presentationUrl: `https://docs.google.com/presentation/d/${presentationId}/edit`,
    slideCount: slideContents.length,
    sections: curriculum.sections,
  };
}

function validateInput_(input) {
  if (!input) throw new Error('Input is empty');

  const title = String(input.title || '').trim();
  const theme = String(input.theme || '').trim();
  const pageCount = Number(input.pageCount || CFG.DEFAULT_PAGE_COUNT);
  const audience = String(input.audience || '').trim() || 'Professional trainers';
  const tone = String(input.tone || '').trim() || 'Practical, safety-first, field-ready';
  const templatePresentationId = String(input.templatePresentationId || '').trim();

  if (!title) throw new Error('Title is required');
  if (!theme) throw new Error('Theme is required');
  if (!Number.isFinite(pageCount)) throw new Error('Page count is invalid');
  if (pageCount < CFG.MIN_PAGE_COUNT || pageCount > CFG.MAX_PAGE_COUNT) {
    throw new Error(`Page count must be between ${CFG.MIN_PAGE_COUNT} and ${CFG.MAX_PAGE_COUNT}`);
  }

  return { title, theme, pageCount, audience, tone, templatePresentationId };
}

function generateCurriculum({ apiKey, model, theme, audience, pageCount }) {
  const prompt = [
    'You are an expert in exercise physiology, rehabilitation, sports science, and elderly fitness training.',
    'Create a lecture curriculum for trainer education.',
    '',
    `Topic: ${theme}`,
    `Audience: ${audience}`,
    `Target slide count: ${pageCount}`,
    '',
    'Requirements:',
    '- Practical education only; avoid abstract explanations.',
    '- Focus on safety and contraindications.',
    '- Include trainer instructions and common mistakes across the curriculum flow.',
    '- Sections should represent lecture topics and be ordered logically for instruction.',
    `- Produce 6-12 sections suitable for approximately ${pageCount} slides.`,
    '',
    'Output JSON only.',
  ].join('\n');

  const schema = {
    type: 'object',
    additionalProperties: false,
    properties: {
      sections: {
        type: 'array',
        minItems: 6,
        maxItems: 12,
        items: { type: 'string', minLength: 3 },
      },
    },
    required: ['sections'],
  };

  const result = callOpenAiJsonWithRetry_({ apiKey, model, prompt, schemaName: 'curriculum_schema', schema });
  if (!Array.isArray(result.sections) || result.sections.length === 0) {
    throw new Error('Curriculum generation failed: sections missing');
  }
  return result;
}

function generateSlideStructure({ apiKey, model, sections, pageCount }) {
  const prompt = [
    'You are designing a professional trainer lecture deck.',
    'Create a slide structure from lecture sections.',
    '',
    `Sections: ${JSON.stringify(sections)}`,
    `Total slides required: ${pageCount}`,
    '',
    'Requirements:',
    '- Each section should have 1-3 slides.',
    '- Sum of slideCount across all sections must equal the total slides required.',
    '- Allocate more slides to safety, coaching workflow, and case practice topics.',
    '',
    'Output JSON only.',
  ].join('\n');

  const schema = {
    type: 'object',
    additionalProperties: false,
    properties: {
      slides: {
        type: 'array',
        minItems: 1,
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            section: { type: 'string', minLength: 3 },
            slideCount: { type: 'integer', minimum: 1, maximum: 3 },
          },
          required: ['section', 'slideCount'],
        },
      },
    },
    required: ['slides'],
  };

  const result = callOpenAiJsonWithRetry_({ apiKey, model, prompt, schemaName: 'slide_structure_schema', schema });
  validateSlideStructure_(result, sections, pageCount);
  return result;
}

function generateSlideContent({ apiKey, model, section, slideIndex, slideCount, audience, tone }) {
  const prompt = [
    'You are a professional trainer educator.',
    'Convert the following lecture topic into one high-quality training slide.',
    '',
    `Topic: ${section}`,
    `Audience: ${audience}`,
    `Deck position: Slide ${slideIndex} of ${slideCount}`,
    `Tone: ${tone}`,
    '',
    'Requirements:',
    '- 3-5 core bullet points focused on practical field execution.',
    '- Include practical trainer instructions in trainer_tips.',
    '- Include elderly safety considerations and contraindications where relevant.',
    '- Include typical trainer mistakes in common_mistakes.',
    '- speaker_notes should be detailed and usable by lecturers directly (4-7 sentences).',
    '- Avoid generic textbook wording; include real-world coaching details.',
    '',
    'Output JSON only in this exact format:',
    '{"title":"","points":[],"trainer_tips":[],"common_mistakes":[],"speaker_notes":""}',
  ].join('\n');

  const schema = {
    type: 'object',
    additionalProperties: false,
    properties: {
      title: { type: 'string', minLength: 3 },
      points: {
        type: 'array',
        minItems: 3,
        maxItems: 5,
        items: { type: 'string', minLength: 3 },
      },
      trainer_tips: {
        type: 'array',
        minItems: 2,
        maxItems: 5,
        items: { type: 'string', minLength: 3 },
      },
      common_mistakes: {
        type: 'array',
        minItems: 2,
        maxItems: 5,
        items: { type: 'string', minLength: 3 },
      },
      speaker_notes: { type: 'string', minLength: 40 },
    },
    required: ['title', 'points', 'trainer_tips', 'common_mistakes', 'speaker_notes'],
  };

  return callOpenAiJsonWithRetry_({ apiKey, model, prompt, schemaName: 'slide_content_schema', schema });
}

function callOpenAiJsonWithRetry_({ apiKey, model, prompt, schemaName, schema }) {
  let lastError = null;

  for (let attempt = 1; attempt <= CFG.RETRY_MAX; attempt++) {
    try {
      const response = callOpenAiRaw_({
        apiKey,
        payload: {
          model,
          input: [{ role: 'user', content: [{ type: 'input_text', text: prompt }] }],
          text: {
            format: {
              type: 'json_schema',
              name: schemaName,
              schema,
              strict: true,
            },
          },
          temperature: 0.3,
        },
      });

      return extractJsonFromResponses_(response);
    } catch (e) {
      lastError = e;
      if (attempt < CFG.RETRY_MAX) {
        const wait = CFG.RETRY_BASE_SLEEP_MS * Math.pow(2, attempt - 1) + Math.floor(Math.random() * 200);
        Utilities.sleep(wait);
      }
    }
  }

  throw new Error(`OpenAI generation failed after ${CFG.RETRY_MAX} attempts: ${lastError && lastError.message}`);
}

function callOpenAiRaw_({ apiKey, payload }) {
  const res = UrlFetchApp.fetch(CFG.OPENAI_ENDPOINT, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${apiKey}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: CFG.REQUEST_TIMEOUT_MS,
  });

  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error(`OpenAI API error HTTP ${code}: ${body}`);
  }

  return safeJsonParse_(body);
}

function extractJsonFromResponses_(json) {
  if (!json) throw new Error('OpenAI response is empty');
  if (json.output_parsed && typeof json.output_parsed === 'object') return json.output_parsed;

  const texts = [];
  if (Array.isArray(json.output)) {
    json.output.forEach((item) => {
      if (!item || !Array.isArray(item.content)) return;
      item.content.forEach((part) => {
        if (part && typeof part.text === 'string') texts.push(part.text);
      });
    });
  }

  if (texts.length > 0) {
    const joined = texts.join('\n').replace(/```json|```/g, '').trim();
    return safeJsonParse_(joined);
  }

  if (typeof json.text === 'string') return safeJsonParse_(json.text);
  throw new Error('Could not extract JSON from OpenAI response');
}

function safeJsonParse_(raw) {
  try {
    return JSON.parse(raw);
  } catch (e) {
    throw new Error(`JSON parse failed: ${e.message}; raw=${String(raw).slice(0, 1000)}`);
  }
}

function validateSlideStructure_(structure, sections, pageCount) {
  if (!structure || !Array.isArray(structure.slides)) {
    throw new Error('Slide structure is invalid');
  }

  const sectionSet = new Set(sections);
  let total = 0;
  structure.slides.forEach((entry, index) => {
    if (!entry || typeof entry.section !== 'string' || !Number.isInteger(entry.slideCount)) {
      throw new Error(`Slide structure item ${index + 1} is invalid`);
    }
    if (!sectionSet.has(entry.section)) {
      throw new Error(`Unknown section in slide structure: ${entry.section}`);
    }
    if (entry.slideCount < 1 || entry.slideCount > 3) {
      throw new Error(`slideCount must be 1-3: ${entry.section}`);
    }
    total += entry.slideCount;
  });

  if (total !== pageCount) {
    throw new Error(`Slide structure total mismatch. expected=${pageCount}, actual=${total}`);
  }
}

function buildSlideBlueprints_(slidesBySection, pageCount) {
  const list = [];
  slidesBySection.forEach((item) => {
    for (let i = 0; i < item.slideCount; i++) {
      list.push({ section: item.section });
    }
  });

  if (list.length !== pageCount) {
    throw new Error(`Slide blueprint count mismatch: expected ${pageCount}, got ${list.length}`);
  }
  return list;
}

function createSlides({ title, templatePresentationId, curriculum, slides }) {
  const presentation = templatePresentationId
    ? copyTemplatePresentation_(templatePresentationId, `${CFG.APP_TITLE_PREFIX}${title}`)
    : SlidesApp.create(`${CFG.APP_TITLE_PREFIX}${title}`);

  cleanupDefaultSlides_(presentation);

  // Agenda slide from curriculum sections
  const agendaSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
  setTitleAndBody_(agendaSlide, `${title} - Curriculum`, curriculum.sections.map((s, i) => `${i + 1}. ${s}`));
  agendaSlide.getNotesPage().getSpeakerNotesShape().getText().setText(
    'Use this slide to explain learning flow, expected outcomes, and where safety checkpoints appear in the training.'
  );

  slides.forEach((content) => {
    const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    const bodyLines = [];
    content.points.forEach((p) => bodyLines.push(p));
    bodyLines.push('');
    bodyLines.push('Trainer tips:');
    content.trainer_tips.forEach((t) => bodyLines.push(`- ${t}`));
    bodyLines.push('');
    bodyLines.push('Common mistakes:');
    content.common_mistakes.forEach((m) => bodyLines.push(`- ${m}`));

    setTitleAndBody_(slide, content.title, bodyLines);
    slide.getNotesPage().getSpeakerNotesShape().getText().setText(content.speaker_notes);
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

function setTitleAndBody_(slide, title, lines) {
  const titlePh = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
  const bodyPh = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);

  const bodyText = (lines || [])
    .map((line) => {
      if (!line) return '';
      if (/^(\d+\.|-|•|Trainer tips:|Common mistakes:)/.test(line)) return line;
      return `• ${line}`;
    })
    .join('\n');

  if (titlePh) titlePh.asShape().getText().setText(title);
  if (bodyPh) {
    bodyPh.asShape().getText().setText(bodyText);
  } else {
    const box = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 60, 110, 610, 320);
    box.getText().setText(bodyText);
  }
}

function getApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) throw new Error('Script property OPENAI_API_KEY is not set');
  return key;
}

function getModel_() {
  return PropertiesService.getScriptProperties().getProperty('OPENAI_MODEL') || CFG.DEFAULT_MODEL;
}
