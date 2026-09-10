/**
 * PTE Practice App — Code.gs
 * Updated behaviour:
 * JSON packages can be imported without a PDF.
 * PDF processing is optional and occurs only when questions explicitly set visualRequired: true.
 * No PDF page images are added when a package has no required visuals.
 * Visual source-page and question-range metadata are saved correctly.
 * Run setupEnvironment() once on a new spreadsheet-bound Apps Script project.
 */

const PTE = {
  SHEETS: {
    TESTS: 'Tests',
    QUESTIONS: 'TestQuestions',
    SETTINGS: 'TestSettings',
    VISUALS: 'VisualAssets',
    READING: 'ReadingProgress',
    RESPONSES: 'UserResponses',
    RESULTS: 'TestResultsSummary'
  },
  VISUAL_FOLDER: 'PTE Visual Assets'
};

const GEMINI_VISUAL_MODEL = 'gemini-3.8-flash';

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiGet(e);
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('PTE Exam')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  return handleApiPost(e);
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleApiGet(e) {
  try {
    const action = e.parameter.action;
    switch(action) {
      case 'getDraftList': return jsonResponse({status: 'success', data: getDraftList()});
      case 'getPublishedTest': return jsonResponse({status: 'success', data: getPublishedTest(e.parameter.testId)});
      case 'getUserHistory': return jsonResponse({status: 'success', data: getUserHistory()});
      case 'getTestAnalysis': return jsonResponse({status: 'success', data: getTestAnalysis(e.parameter.testId)});
      case 'getOverallAnalysis': return jsonResponse(getOverallAnalysis());
      default: return jsonResponse({status: 'error', message: 'Unknown action'});
    }
  } catch(error) {
    return jsonResponse({status: 'error', message: error.message});
  }
}

function handleApiPost(e) {
  try {
    let payload = {};
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }
    const action = payload.action;
    switch(action) {
      case 'getDraftList': return jsonResponse({status: 'success', data: getDraftList()});
      case 'getPublishedTest': return jsonResponse({status: 'success', data: getPublishedTest(payload.testId)});
      case 'getUserHistory': return jsonResponse({status: 'success', data: getUserHistory()});
      case 'getTestAnalysis': return jsonResponse({status: 'success', data: getTestAnalysis(payload.testId)});
      case 'getOverallAnalysis': return jsonResponse(getOverallAnalysis());
      case 'importStructuredTestPackage': return jsonResponse(importStructuredTestPackage(payload));
      case 'massUploadExams': return jsonResponse(massUploadExams(payload));
      case 'saveAdminDraft': return jsonResponse(saveAdminDraft(payload));
      case 'publishTest': return jsonResponse(publishTest(payload));
      case 'deleteDraft': return jsonResponse(deleteDraft(payload.testId));
      case 'submitBlockAnswers': return jsonResponse(submitBlockAnswers(payload));
      case 'recordReadingCompletion': return jsonResponse(recordReadingCompletion(payload));
      case 'assessPteSpeakingWithGemini': return jsonResponse(assessPteSpeakingWithGemini(payload));
      case 'checkAnswer': return jsonResponse(checkAnswer(payload));
      case 'generateSamplePteTest': return jsonResponse(generateSamplePteTest());
      default: return jsonResponse({status: 'error', message: 'Unknown action (' + action + ')'});
    }
  } catch(error) {
    return jsonResponse({status: 'error', message: error.message});
  }
}

function setupEnvironment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  resetSheet_(ss, PTE.SHEETS.TESTS, [ 'TestID', 'Title', 'Status', 'CreatedAt', 'PublishedAt', 'SourceFileName' ]);
  resetSheet_(ss, PTE.SHEETS.QUESTIONS, questionHeaders_());
  resetSheet_(ss, PTE.SHEETS.SETTINGS, [ 'TestID', 'RequireReadingCompletion', 'ImmediateFeedback', 'ShowRunningScore', 'EnablePostBlockReview', 'Published', 'UpdatedAt' ]);
  resetSheet_(ss, PTE.SHEETS.VISUALS, [ 'AssetID', 'TestID', 'SourcePage', 'QuestionStart', 'QuestionEnd', 'AssetType', 'DriveFileID', 'AltText', 'CreatedAt' ]);
  ensureSheet_(ss, PTE.SHEETS.READING, [ 'Timestamp', 'UserEmail', 'TestID', 'BlockID', 'ReadingOpenedAt', 'ReadingConfirmedAt', 'ReachedEnd', 'QuestionStageUnlocked' ]);
  ensureSheet_(ss, PTE.SHEETS.RESPONSES, [ 'Timestamp', 'UserEmail', 'TestID', 'BlockID', 'QuestionNumber', 'UserAnswer', 'IsCorrect', 'TimeSeconds' ]);
  ensureSheet_(ss, PTE.SHEETS.RESULTS, [ 'Timestamp', 'UserEmail', 'TestID', 'Block1', 'Block2', 'Block3', 'Block4', 'TotalOutOf50', 'BandScore', 'Band7Achieved' ]);

  getOrCreateVisualFolder_();
  return { status: 'success', message: 'Environment created. Upload a valid JSON package in Admin mode.' };
}

function questionHeaders_() {
  return [ 'TestID', 'QuestionNumber', 'BlockID', 'QuestionType', 'Prompt', 'Instruction', 'OptionsJSON', 'AcceptedAnswersJSON', 'Explanation', 'LookoutTip', 'VisualAssetID', 'VisualSourcePage', 'VisualQuestionStart', 'VisualQuestionEnd', 'VisualRequired', 'VisualAltText', 'RequiresReview', 'ReviewNote', 'MinWordLimit', 'MaxWordLimit', 'UpdatedAt' ];
}

function resetSheet_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clear();
  sheet.appendRow(headers);
  styleHeader_(sheet, headers.length);
  return sheet;
}

function ensureSheet_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    styleHeader_(sheet, headers.length);
  }
  return sheet;
}

function styleHeader_(sheet, count) {
  sheet.getRange(1, 1, 1, count)
    .setFontWeight('bold')
    .setFontColor('#334155')
    .setBackground('#F1F5F9');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, count);
}

function getOrCreateVisualFolder_() {
  const folders = DriveApp.getFoldersByName(PTE.VISUAL_FOLDER);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(PTE.VISUAL_FOLDER);
}

function massUploadExams(payload) {
  if (!payload || !Array.isArray(payload.packages) || payload.packages.length === 0) {
    throw new Error('No packages were supplied.');
  }
  let results = [];
  payload.packages.forEach((pkg, index) => {
    try {
      if (index >= 10) return; // limit to 10
      let draft;
      if (pkg.isPdf) {
         draft = generatePackageFromPdf_(pkg.data, pkg.fileName || `PTE_Test_${index+1}.pdf`);
      } else {
         let imported = typeof pkg.data === 'string' ? JSON.parse(pkg.data) : pkg.data;
         draft = buildDraftFromPackage_(imported, pkg.fileName || `PTE_Test_${index+1}.json`);
      }

      let check = validateDraft_(draft, false);
      draft.status = check.valid ? 'DRAFT — READY FOR REVIEW' : 'DRAFT — REVIEW REQUIRED';
      saveNewDraft_(draft);
      results.push({status: 'success', testId: draft.testId});
    } catch(err) {
      results.push({status: 'error', message: err.message, fileName: pkg.fileName});
    }
  });
  return {status: 'success', results: results};
}

function importStructuredTestPackage(payload) {
  if (!payload || !payload.jsonText) {
    throw new Error('No JSON package was supplied.');
  }
  let imported;
  try {
    imported = JSON.parse(String(payload.jsonText));
  } catch (error) {
    throw new Error('The uploaded file is not valid JSON: ' + error.message);
  }

  const draft = buildDraftFromPackage_(imported, payload.fileName || 'PTE_Test_Package.json');
  const check = validateDraft_(draft, false);
  draft.status = check.valid ? 'DRAFT — READY FOR REVIEW' : 'DRAFT — REVIEW REQUIRED';

  saveNewDraft_(draft);
  return getAdminDraft(draft.testId);
}

function buildDraftFromPackage_(packageData, fileName) {
  const packageId = String(packageData.testId || '').trim();
  const testId = (!packageId || packageId === 'REPLACE_WITH_TEST_ID') ? 'TEST_' + Date.now() : packageId + '_' + Date.now();

  const blocks = Array.isArray(packageData.blocks) ? packageData.blocks : [];
  const passages = {};
  const suppliedQuestions = [];

  blocks.forEach((block, index) => {
    const blockId = Number(block.blockId || index + 1);
    passages['p' + blockId + 'Title'] = String(block.passageTitle || 'Passage ' + blockId);
    passages['p' + blockId + 'Text'] = tidyPassageText_(block.passageText || '');

    const blockQuestions = Array.isArray(block.questions) ? block.questions : [];
    blockQuestions.forEach(question => {
      suppliedQuestions.push(Object.assign({}, question, {blockId: blockId}));
    });
  });

  return {
    testId: testId,
    title: String(packageData.title || fileName.replace(/\.json$/i, '') || 'PTE Exam'),
    sourceFileName: String(packageData.sourceFileName || fileName),
    status: 'DRAFT — REVIEW REQUIRED',
    passages: passages,
    settings: {
      requireReadingCompletion: !packageData.settings || packageData.settings.requireReadingCompletion !== false,
      immediateFeedback: !packageData.settings || packageData.settings.immediateFeedback !== false,
      showRunningScore: !packageData.settings || packageData.settings.showRunningScore !== false,
      enablePostBlockReview: !packageData.settings || packageData.settings.enablePostBlockReview !== false,
      published: false
    },
    questions: normaliseQuestions_(suppliedQuestions)
  };
}

function tidyPassageText_(text) {
  return String(text || '')
    .replace(/\n/g, '\n')
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function normaliseQuestions_(suppliedQuestions) {
  const byNumber = {};
  (suppliedQuestions || []).forEach(question => {
    const number = Number(question.qNum);
    if (number >= 1 && number <= 50 && !byNumber[number]) {
      byNumber[number] = question;
    }
  });

  const output = [];
  for (let number = 1; number <= 50; number++) {
    const q = byNumber[number] || {};
    const visual = q.visual && typeof q.visual === 'object' ? q.visual : null;
    const answers = Array.isArray(q.acceptedAnswers) ? q.acceptedAnswers.map(value => String(value).trim()).filter(Boolean) : [];
    const prompt = String(q.prompt || '').trim();
    const type = validQuestionType_(q.type || defaultQuestionType_(number));
    const visualRequired = q.visualRequired === true || (visual && visual.required === true);

    output.push({
      qNum: number,
      blockId: Number(q.blockId || expectedBlock_(number)),
      type: type,
      prompt: prompt || defaultPrompt_(number),
      instruction: String(q.instruction || '').trim() || defaultInstruction_(type),
      options: Array.isArray(q.options) ? q.options.map(value => String(value).trim()).filter(Boolean) : [],
      acceptedAnswers: answers.length ? answers : ['REVIEW REQUIRED'],
      explanation: String(q.explanation || '').trim() || 'REVIEW REQUIRED: verify the explanation before publishing.',
      lookoutTips: String(q.lookoutTips || '').trim() || 'REVIEW REQUIRED: add a helpful PTE strategy tip.',
      visualAssetId: String(q.visualAssetId || ''),
      visualSourcePage: visual && Number(visual.sourcePage)
        ? Number(visual.sourcePage)
        : Number(q.visualSourcePage || 0),
      visualQuestionStart: Number(q.visualQuestionStart || (visual && visual.questionStart) || 0),
      visualQuestionEnd: Number(q.visualQuestionEnd || (visual && visual.questionEnd) || 0),
      visualAltText: visual && visual.altText ? String(visual.altText) : String(q.visualAltText || ''),
      visualRequired: visualRequired,
      requiresReview: q.requiresReview === true || !byNumber[number] || !prompt || !answers.length ||
        (visualRequired && !String(q.visualAssetId || '').trim()),
      reviewNote: String(q.reviewNote || ''),
      MinWordLimit: Number(q.MinWordLimit || (String(q.instruction || '').toLowerCase().includes('summarize') ? 5 : String(q.instruction || '').toLowerCase().includes('essay') ? 200 : 0)),
      MaxWordLimit: Number(q.MaxWordLimit || (String(q.instruction || '').toLowerCase().includes('summarize') ? 75 : String(q.instruction || '').toLowerCase().includes('essay') ? 300 : Infinity))
    });
  }
  return output;
}

function validQuestionType_(type) {
  const allowed = [ 'text', 'diagram_label', 'table_completion', 'multiple_choice', 'true_false_not_given', 'matching', 'matching_headings', 'matching_paragraphs', 'speaking' ];
  const candidate = String(type || '').toLowerCase().trim();
  return allowed.indexOf(candidate) >= 0 ? candidate : 'text';
}

function expectedBlock_(number) {
  return number <= 15 ? 1 : number <= 20 ? 2 : number <= 35 ? 3 : 4;
}

function defaultQuestionType_(number) {
  if (number <= 15) return 'speaking';
  if (number <= 20) return 'text'; // writing
  if (number >= 21 && number <= 35) return 'multiple_choice'; // reading
  if (number >= 36 && number <= 50) return 'multiple_choice'; // listening
  return 'text';
}

function defaultPrompt_(number) {
  return 'REVIEW REQUIRED: add Question ' + number + '.';
}

function defaultInstruction_(type) {
  const map = {
    diagram_label: 'Choose NO MORE THAN TWO WORDS from the reading passage. Use the diagram and passage together.',
    table_completion: 'Choose NO MORE THAN TWO WORDS AND/OR A NUMBER from the passage.',
    true_false_not_given: 'Choose True, False, or Not Given.',
    matching: 'Match the question with the best option.',
    matching_headings: 'Choose the correct heading.',
    matching_paragraphs: 'Choose the paragraph containing this information.',
    multiple_choice: 'Choose the correct answer.',
    text: 'Write the required answer from the passage.',
    speaking: 'Read the text aloud or answer the question verbally.'
  };
  return map[type] || map.text;
}

function validateDraft_(draft, strict) {
  const issues = [];
  const passages = draft.passages || {};

  [1, 2, 3, 4].forEach(block => {
    if (!String(passages['p' + block + 'Text'] || '').trim()) {
      issues.push('Passage ' + block + ' is empty.');
    }
  });

  if (!Array.isArray(draft.questions) || draft.questions.length !== 50) {
    issues.push('The test must contain exactly 50 questions.');
  }

  for (let i = 0; i < 50; i++) {
    const question = draft.questions && draft.questions[i];
    const number = i + 1;

    if (!question || Number(question.qNum) !== number) {
      issues.push('Question ' + number + ' is missing or incorrectly numbered.');
      continue;
    }

    if (!String(question.prompt || '').trim()) {
      issues.push('Question ' + number + ' has a blank prompt.');
    }

    if (!Array.isArray(question.acceptedAnswers) || !question.acceptedAnswers.length) {
      issues.push('Question ' + number + ' has no accepted answer.');
    }

    if (strict && question.visualRequired && !String(question.visualAssetId || '').trim()) {
      issues.push('Question ' + number + ' requires an existing visual, but no image has been uploaded.');
    }

    if (strict && String(question.prompt || '').indexOf('REVIEW REQUIRED') >= 0) {
      issues.push('Question ' + number + ' prompt still requires review.');
    }

    if (strict && (question.acceptedAnswers || []).indexOf('REVIEW REQUIRED') >= 0) {
      issues.push('Question ' + number + ' answer still requires review.');
    }
  }

  if (strict && issues.length) {
    throw new Error('This test cannot be published until these issues are fixed:\n\n' + issues.join('\n'));
  }

  return {valid: issues.length === 0, issues: issues};
}

function saveNewDraft_(draft) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tests = ensureSheet_(ss, PTE.SHEETS.TESTS, [ 'TestID', 'Title', 'Status', 'CreatedAt', 'PublishedAt', 'SourceFileName' ]);
  const questions = ensureSheet_(ss, PTE.SHEETS.QUESTIONS, questionHeaders_());
  const settings = ensureSheet_(ss, PTE.SHEETS.SETTINGS, [ 'TestID', 'RequireReadingCompletion', 'ImmediateFeedback', 'ShowRunningScore', 'EnablePostBlockReview', 'Published', 'UpdatedAt' ]);

  tests.appendRow([ draft.testId, draft.title, draft.status, new Date().toISOString(), '', draft.sourceFileName ]);

  const questionRows = draft.questions.map(question => questionRow_(draft.testId, question));
  if (questionRows.length) {
    questions.getRange(questions.getLastRow() + 1, 1, questionRows.length, questionRows[0].length)
      .setValues(questionRows);
  }

  settings.appendRow([ draft.testId, draft.settings.requireReadingCompletion, draft.settings.immediateFeedback, draft.settings.showRunningScore, draft.settings.enablePostBlockReview, false, new Date().toISOString() ]);

  PropertiesService.getDocumentProperties()
    .setProperty('PASSAGES_' + draft.testId, JSON.stringify(draft.passages));
}

function questionRow_(testId, question) {
  return [
    testId, question.qNum, question.blockId, question.type, question.prompt, question.instruction,
    JSON.stringify(question.options || []), JSON.stringify(question.acceptedAnswers || []),
    question.explanation || '', question.lookoutTips || '', question.visualAssetId || '',
    question.visualSourcePage || '', question.visualQuestionStart || '', question.visualQuestionEnd || '',
    question.visualRequired ? 'YES' : 'NO', question.visualAltText || '', question.requiresReview ? 'YES' : 'NO',
    question.reviewNote || '', question.MinWordLimit || 0, question.MaxWordLimit || 0, new Date().toISOString()
  ];
}


function uploadQuestionVisual(payload) {
  if (!payload || !payload.testId || !payload.base64Data) {
    throw new Error('Test ID and image data are required.');
  }

  const questionStart = Number(payload.questionStart);
  const questionEnd = Number(payload.questionEnd);

  if (!Number.isInteger(questionStart) || !Number.isInteger(questionEnd) || questionStart < 1 || questionEnd > 50 || questionStart > questionEnd) {
    throw new Error('Invalid question range for the visual asset.');
  }

  const blob = Utilities.newBlob( Utilities.base64Decode(payload.base64Data), payload.mimeType || 'image/png', payload.fileName || 'pte-visual.png' );

  const file = getOrCreateVisualFolder_().createFile(blob);
  const assetId = 'ASSET_' + Date.now() + '_' + Math.floor(Math.random() * 100000);
  const sourcePage = Number(payload.sourcePage || 0);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const visualSheet = ensureSheet_(ss, PTE.SHEETS.VISUALS, [ 'AssetID', 'TestID', 'SourcePage', 'QuestionStart', 'QuestionEnd', 'AssetType', 'DriveFileID', 'AltText', 'CreatedAt' ]);

  visualSheet.appendRow([ assetId, payload.testId, sourcePage, questionStart, questionEnd, 'admin_uploaded_image', file.getId(), payload.altText || 'Visual reference for Questions ' + questionStart + '–' + questionEnd, new Date().toISOString() ]);

  attachVisualToQuestionRange_(payload.testId, assetId, questionStart, questionEnd, sourcePage);

  return { status: 'success', assetId: assetId, driveFileId: file.getId(), questionStart: questionStart, questionEnd: questionEnd, sourcePage: sourcePage, message: 'Visual attached to Questions ' + questionStart + '–' + questionEnd + '.' };
}

function attachVisualToQuestionRange_(testId, assetId, questionStart, questionEnd, sourcePage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PTE.SHEETS.QUESTIONS);
  if (!sheet) throw new Error('TestQuestions sheet was not found.');

  const values = sheet.getDataRange().getValues();
  values.slice(1).forEach((row, index) => {
    const rowTestId = String(row[0]);
    const questionNumber = Number(row[1]);

    if (rowTestId === String(testId) && questionNumber >= questionStart && questionNumber <= questionEnd) {
      const rowNumber = index + 2;
      sheet.getRange(rowNumber, 11, 1, 4).setValues([[
        assetId,
        Number(sourcePage || 0),
        questionStart,
        questionEnd
      ]]);
    }
  });
}

function uploadVisualAsset(payload) { return uploadQuestionVisual(payload); }

function getDraftList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PTE.SHEETS.TESTS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getDataRange().getValues().slice(1).map(row => ({
    testId: row[0], title: row[1], status: row[2], createdAt: row[3], publishedAt: row[4]
  })).reverse();
}

function getAdminDraft(testId) {
  const record = getTestRecord_(testId);
  if (!record) throw new Error('Test draft was not found.');

  return { test: record, passages: getPassages_(testId), questions: getQuestions_(testId), settings: getSettings_(testId), visuals: getVisuals_(testId) };
}

function getTestRecord_(testId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PTE.SHEETS.TESTS);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const row = sheet.getDataRange().getValues().slice(1)
    .find(item => String(item[0]) === String(testId));

  return row ? { testId: row[0], title: row[1], status: row[2], createdAt: row[3], publishedAt: row[4], sourceFileName: row[5] } : null;
}

function getQuestions_(testId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PTE.SHEETS.QUESTIONS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getDataRange().getValues().slice(1)
    .filter(row => String(row[0]) === String(testId))
    .map(row => ({
      qNum: Number(row[1]), blockId: Number(row[2]), type: row[3], prompt: row[4], instruction: row[5],
      options: jsonArray_(row[6]), acceptedAnswers: jsonArray_(row[7]), explanation: row[8],
      lookoutTips: row[9], visualAssetId: row[10], visualSourcePage: Number(row[11] || 0),
      visualQuestionStart: Number(row[12] || 0), visualQuestionEnd: Number(row[13] || 0),
      visualRequired: String(row[14]).toUpperCase() === 'YES', visualAltText: row[15] || '',
      requiresReview: String(row[16]).toUpperCase() === 'YES', reviewNote: row[17] || '',
      MinWordLimit: Number(row[18] || 0), MaxWordLimit: Number(row[19] || 0)
    }))
    .sort((a, b) => a.qNum - b.qNum);
}

function jsonArray_(value) {
  try {
    const result = JSON.parse(value || '[]');
    return Array.isArray(result) ? result : [];
  } catch (error) {
    return [];
  }
}

function getSettings_(testId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PTE.SHEETS.SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return defaultSettings_();

  const row = sheet.getDataRange().getValues().slice(1)
    .find(item => String(item[0]) === String(testId));

  if (!row) return defaultSettings_();

  return { requireReadingCompletion: asBoolean_(row[1]), immediateFeedback: asBoolean_(row[2]), showRunningScore: asBoolean_(row[3]), enablePostBlockReview: asBoolean_(row[4]), published: asBoolean_(row[5]) };
}

function defaultSettings_() {
  return { requireReadingCompletion: true, immediateFeedback: true, showRunningScore: true, enablePostBlockReview: true, published: false };
}

function asBoolean_(value) {
  return value === true || String(value).toUpperCase() === 'TRUE' || String(value).toUpperCase() === 'YES';
}

function getPassages_(testId) {
  try {
    return JSON.parse(PropertiesService.getDocumentProperties().getProperty('PASSAGES_' + testId) || '{}');
  } catch (error) {
    return {};
  }
}

function getVisuals_(testId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PTE.SHEETS.VISUALS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getDataRange().getValues().slice(1)
    .filter(row => !testId || String(row[1]) === String(testId))
    .map(row => ({ assetId: row[0], testId: row[1], sourcePage: Number(row[2]), questionStart: Number(row[3]), questionEnd: Number(row[4]), assetType: row[5], driveFileId: row[6], altText: row[7] }));
}

function deleteTest(testId) {
  if (!testId) throw new Error('Test ID is required.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tests = ss.getSheetByName(PTE.SHEETS.TESTS);
  if (!tests || tests.getLastRow() < 2) throw new Error('Test not found.');

  const testRows = tests.getDataRange().getValues();
  const record = testRows.slice(1).find(row => String(row[0]) === String(testId));
  if (!record) throw new Error('Test not found.');

  const visuals = getVisuals_(testId);

  deleteRowsByTestId_(tests, testId, 0);
  deleteRowsByTestId_(ss.getSheetByName(PTE.SHEETS.QUESTIONS), testId, 0);
  deleteRowsByTestId_(ss.getSheetByName(PTE.SHEETS.SETTINGS), testId, 0);
  deleteRowsByTestId_(ss.getSheetByName(PTE.SHEETS.VISUALS), testId, 1);
  deleteRowsByTestId_(ss.getSheetByName(PTE.SHEETS.READING), testId, 2);
  deleteRowsByTestId_(ss.getSheetByName(PTE.SHEETS.RESPONSES), testId, 2);
  deleteRowsByTestId_(ss.getSheetByName(PTE.SHEETS.RESULTS), testId, 2);

  visuals.forEach(asset => {
    if (!asset.driveFileId) return;
    try {
      DriveApp.getFileById(asset.driveFileId).setTrashed(true);
    } catch (error) {
      Logger.log('Unable to trash visual file ' + asset.driveFileId + ': ' + error);
    }
  });

  PropertiesService.getDocumentProperties().deleteProperty('PASSAGES_' + testId);

  return {status: 'success', message: 'Test, learner records for this test, and attached visual files deleted.'};
}

function deleteDraft(testId) {
  return deleteTest(testId);
}

function deleteRowsByTestId_(sheet, testId, testIdColumnZeroBased) {
  if (!sheet || sheet.getLastRow() < 2) return;

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const remaining = values.slice(1).filter(row => String(row[testIdColumnZeroBased]) !== String(testId));

  sheet.clear();
  sheet.appendRow(headers);
  if (remaining.length) {
    sheet.getRange(2, 1, remaining.length, headers.length).setValues(remaining);
  }
  styleHeader_(sheet, headers.length);
}

function saveAdminDraft(payload) {
  if (!payload || !payload.testId) throw new Error('Test ID is missing.');

  const existingByNumber = {};
  getQuestions_(payload.testId).forEach(question => { existingByNumber[question.qNum] = question; });

  const mergedQuestions = (payload.questions || []).map(question => {
    const existing = existingByNumber[Number(question.qNum)] || {};
    return Object.assign({}, existing, question, {
      visualQuestionStart: Number(question.visualQuestionStart || existing.visualQuestionStart || 0),
      visualQuestionEnd: Number(question.visualQuestionEnd || existing.visualQuestionEnd || 0),
      requiresReview: question.requiresReview === true || existing.requiresReview === true,
      reviewNote: question.reviewNote !== undefined ? question.reviewNote : (existing.reviewNote || '')
    });
  });

  const draft = { testId: payload.testId, passages: payload.passages || {}, questions: normaliseQuestions_(mergedQuestions), settings: Object.assign(defaultSettings_(), payload.settings || {}) };

  const validation = validateDraft_(draft, false);
  replaceQuestions_(draft.testId, draft.questions);
  replaceSettings_(draft.testId, draft.settings, false);
  PropertiesService.getDocumentProperties().setProperty('PASSAGES_' + draft.testId, JSON.stringify(draft.passages));
  setTestStatus_(draft.testId, validation.valid ? 'DRAFT — READY FOR REVIEW' : 'DRAFT — REVIEW REQUIRED', '');

  return {status: 'success', message: 'Draft saved.'};
}

function publishTest(payload) {
  if (!payload || !payload.testId) throw new Error('Test ID is missing.');

  // If payload does not contain questions, fetch from the database
  const existingQuestions = getQuestions_(payload.testId);
  const sourceQuestions = (payload.questions && payload.questions.length > 0) ? payload.questions : existingQuestions;

  const existingByNumber = {};
  existingQuestions.forEach(question => { existingByNumber[question.qNum] = question; });

  const mergedQuestions = sourceQuestions.map(question => {
    const existing = existingByNumber[Number(question.qNum)] || {};
    return Object.assign({}, existing, question, {
      visualQuestionStart: Number(question.visualQuestionStart || existing.visualQuestionStart || 0),
      visualQuestionEnd: Number(question.visualQuestionEnd || existing.visualQuestionEnd || 0),
      requiresReview: question.requiresReview === true || existing.requiresReview === true,
      reviewNote: question.reviewNote !== undefined ? question.reviewNote : (existing.reviewNote || '')
    });
  });

  // If payload does not contain passages, fetch from the database
  const passages = payload.passages && Object.keys(payload.passages).length > 0 ? payload.passages : getPassages_(payload.testId);

  // If payload does not contain settings, fetch from the database
  const settings = payload.settings ? Object.assign(getSettings_(payload.testId), payload.settings) : getSettings_(payload.testId);

  const draft = { testId: payload.testId, passages: passages, questions: normaliseQuestions_(mergedQuestions), settings: settings };

  validateDraft_(draft, true);
  draft.questions.forEach(question => { question.requiresReview = false; });
  draft.settings.published = true;

  replaceQuestions_(draft.testId, draft.questions);
  replaceSettings_(draft.testId, draft.settings, true);
  PropertiesService.getDocumentProperties().setProperty('PASSAGES_' + draft.testId, JSON.stringify(draft.passages));
  setTestStatus_(draft.testId, 'PUBLISHED', new Date().toISOString());

  return {status: 'success', message: 'Test published successfully.'};
}

function replaceQuestions_(testId, questions) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PTE.SHEETS.QUESTIONS);
  if (!sheet) throw new Error('TestQuestions sheet was not found. Run setupEnvironment first.');

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const retained = values.slice(1).filter(row => String(row[0]) !== String(testId));

  sheet.clear();
  sheet.appendRow(headers);
  if (retained.length) sheet.getRange(2, 1, retained.length, headers.length).setValues(retained);

  const rows = questions.slice().sort((a, b) => a.qNum - b.qNum).map(question => questionRow_(testId, question));
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }
  styleHeader_(sheet, headers.length);
}

function replaceSettings_(testId, settings, published) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PTE.SHEETS.SETTINGS);
  if (!sheet) throw new Error('TestSettings sheet was not found. Run setupEnvironment first.');

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const retained = values.slice(1).filter(row => String(row[0]) !== String(testId));

  sheet.clear();
  sheet.appendRow(headers);
  if (retained.length) sheet.getRange(2, 1, retained.length, headers.length).setValues(retained);

  sheet.appendRow([ testId, asBoolean_(settings.requireReadingCompletion), asBoolean_(settings.immediateFeedback), asBoolean_(settings.showRunningScore), asBoolean_(settings.enablePostBlockReview), Boolean(published), new Date().toISOString() ]);

  styleHeader_(sheet, headers.length);
}

function setTestStatus_(testId, status, publishedAt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PTE.SHEETS.TESTS);
  if (!sheet) throw new Error('Tests sheet was not found.');

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(testId)) {
      sheet.getRange(i + 1, 3).setValue(status);
      sheet.getRange(i + 1, 5).setValue(publishedAt || '');
      return;
    }
  }
  throw new Error('Test record was not found.');
}

function getPublishedTest(testId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.getSheetByName(PTE.SHEETS.TESTS);
  if (!testSheet || testSheet.getLastRow() < 2) return null;

  const record = testSheet.getDataRange().getValues().slice(1).reverse()
    .find(row => String(row[2]) === 'PUBLISHED' && (!testId || String(row[0]) === String(testId)));

  if (!record) return null;

  const id = record[0];
  const visuals = getVisuals_(id);
  const questions = getQuestions_(id).map(question => {
    const asset = visuals.find(visual => String(visual.assetId) === String(question.visualAssetId));
    if (asset) {
      question.visual = { required: true, driveFileId: asset.driveFileId, altText: asset.altText, sourcePage: asset.sourcePage };
    }
    return question;
  });

  const passages = getPassages_(id);

  return {
    testId: id,
    title: record[1],
    settings: getSettings_(id),
    blocks: [
      buildBlock_(1, "Speaking", passages.p1Title || "Passage 1", passages.p1Text || "", questions, 1, 15),
      buildBlock_(2, "Writing", passages.p2Title || "Passage 2", passages.p2Text || "", questions, 16, 20),
      buildBlock_(3, "Reading", passages.p3Title || "Passage 3", passages.p3Text || "", questions, 21, 35),
      buildBlock_(4, "Listening", passages.p4Title || "Passage 4", passages.p4Text || "", questions, 36, 50)
    ]
  };
}

function buildBlock_(blockId, title, passageTitle, passageText, questions, firstQuestion, lastQuestion) {
  return {
    blockId: blockId,
    title: title,
    timeLimitMinutes: 20,
    passageTitle: passageTitle,
    passageText: passageText,
    questions: questions.filter(question => question.qNum >= firstQuestion && question.qNum <= lastQuestion)
  };
}

function recordReadingCompletion(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheet_(ss, PTE.SHEETS.READING, [ 'Timestamp', 'UserEmail', 'TestID', 'BlockID', 'ReadingOpenedAt', 'ReadingConfirmedAt', 'ReachedEnd', 'QuestionStageUnlocked' ]);

  sheet.appendRow([ new Date(), Session.getActiveUser().getEmail() || 'anonymous_candidate', payload.testId, Number(payload.blockId), payload.readingOpenedAt || '', payload.readingConfirmedAt || new Date().toISOString(), payload.reachedEnd ? 'YES' : 'NO', payload.reachedEnd ? 'YES' : 'NO' ]);

  return {status: 'success'};
}

function checkAnswer(payload) {
  const question = getQuestions_(payload.testId).find(item => item.qNum === Number(payload.qNum));

  if (!question) throw new Error('Question not found.');

  return {
    isCorrect: answersMatch_(payload.answer, question.acceptedAnswers),
    acceptedAnswers: question.acceptedAnswers,
    explanation: question.explanation,
    lookoutTips: question.lookoutTips,
    strategy: answerStrategy_(question)
  };
}

function assessPteSpeakingWithGemini(payload) {
  if (!payload || !payload.audioBase64 || !payload.prompt) {
    throw new Error('A recorded answer and speaking prompt are required.');
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Set GEMINI_API_KEY in Apps Script Project Settings > Script properties.');

  const model = PropertiesService.getScriptProperties().getProperty('GEMINI_VISUAL_MODEL') || GEMINI_VISUAL_MODEL;
  const instruction = 'Assess this PTE speaking practice response. Return JSON only: {"overallScore":0,"fluency":0,"pronunciation":0,"content":0,"transcript":"","feedback":"","tactics":""}. Scores are integer practice estimates from 10 to 90. Give concrete, supportive feedback on pace, pauses, clarity, stress, content coverage and one specific next attempt drill. Prompt: ' + String(payload.prompt);

  const request = {
    contents: [{
      parts: [
        {inline_data: {mime_type: payload.mimeType || 'audio/webm', data: String(payload.audioBase64)}},
        {text: instruction}
      ]
    }],
    generationConfig: {responseMimeType: 'application/json', temperature: 0.2}
  };

  const response = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(apiKey),
    {method: 'post', contentType: 'application/json', payload: JSON.stringify(request), muteHttpExceptions: true}
  );

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error('Gemini speaking assessment failed: ' + response.getContentText().slice(0, 500));
  }

  try {
    const body = JSON.parse(response.getContentText());
    const candidate = body.candidates && body.candidates[0];
    const parts = candidate && candidate.content && candidate.content.parts;
    const assessment = JSON.parse((parts || []).map(part => part.text || '').join(''));
    const score = Math.max(10, Math.min(90, Number(assessment.overallScore) || 10));

    return {
      isCorrect: score >= 58,
      overallScore: score,
      acceptedAnswers: ['Practice estimate: ' + score + ' / 90'],
      explanation: String(assessment.feedback || ''),
      lookoutTips: String(assessment.tactics || ''),
      strategy: 'Use the transcript to identify one pronunciation issue and one fluency issue. Record again once, concentrating on the single drill rather than trying to change everything at once.',
      transcript: String(assessment.transcript || '')
    };
  } catch (error) {
    throw new Error('Gemini did not return a usable speaking assessment: ' + error.message);
  }
}

function submitBlockAnswers(payload) {
  const email = Session.getActiveUser().getEmail() || 'anonymous_candidate';
  const test = getPublishedTest(payload.testId);
  if (!test) throw new Error('Published test not found.');

  const block = test.blocks.find(item => item.blockId === Number(payload.blockId));
  if (!block) throw new Error('Block not found.');

  let score = 0;
  const responseRows = [];

  const results = block.questions.map(question => {
    const answer = String((payload.answers || {})[question.qNum] || '');
    const speakingFeedback = (payload.feedback || {})[question.qNum];
    const correct = question.type === 'speaking' && speakingFeedback ? Boolean(speakingFeedback.isCorrect) : answersMatch_(answer, question.acceptedAnswers);

    if (correct) score++;

    responseRows.push([
      new Date(), email, test.testId, block.blockId, question.qNum,
      answer, correct ? 'YES' : 'NO', Number((payload.times || {})[question.qNum] || 0)
    ]);

    return {
      qNum: question.qNum,
      userAnswer: answer,
      isCorrect: correct,
      acceptedAnswers: question.type === 'speaking' && speakingFeedback ? speakingFeedback.acceptedAnswers : question.acceptedAnswers,
      explanation: question.type === 'speaking' && speakingFeedback ? speakingFeedback.explanation : question.explanation,
      lookoutTips: question.type === 'speaking' && speakingFeedback ? speakingFeedback.lookoutTips : question.lookoutTips,
      strategy: question.type === 'speaking' && speakingFeedback ? speakingFeedback.strategy : answerStrategy_(question)
    };
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = ensureSheet_(ss, PTE.SHEETS.RESPONSES, [ 'Timestamp', 'UserEmail', 'TestID', 'BlockID', 'QuestionNumber', 'UserAnswer', 'IsCorrect', 'TimeSeconds' ]);

  if (responseRows.length) {
    responseSheet.getRange(responseSheet.getLastRow() + 1, 1, responseRows.length, responseRows[0].length).setValues(responseRows);
  }

  return { blockId: block.blockId, score: score, total: block.questions.length, results: results, finalSummary: block.blockId === 4 ? saveFinalSummary_(email, test.testId) : null };
}

function answerStrategy_(question) {
  const type = String(question.type || 'text');
  const strategies = {
    speaking: 'Focus on oral fluency and pronunciation. Do not hesitate or self-correct, and try to speak at a natural pace with clear intonation. If you make a mistake, keep going.',
    diagram_label: 'Step 1: Scan for visual clues or keywords in the image. Step 2: Locate corresponding descriptions in the passage. Step 3: Use exact words from the text.',
    table_completion: 'Step 1: Use row/column headers as anchors. Step 2: Predict grammar and scan the text. Step 3: Copy exact wording respecting word limits.',
    true_false_not_given: 'Step 1: Locate the claim in the text. Step 2: Ensure a full match for True, contradiction for False, and lack of information for Not Given.',
    matching: 'Step 1: Identify unique stems. Step 2: Eliminate options with grammar mismatches. Step 3: Confirm the match with text evidence.',
    matching_headings: 'Step 1: Read topic sentences. Step 2: Ignore specific examples. Step 3: Choose the broad heading covering the entire section.',
    matching_paragraphs: 'Step 1: Use distinctive search terms. Step 2: Scan for ideas, not just words. Step 3: Check context for a full answer.',
    multiple_choice: 'Step 1: Understand the core question. Step 2: Find evidence before looking at options. Step 3: Eliminate options with extreme or unsupported words.',
    text: 'Step 1: Identify key nouns. Step 2: Read around the matching sentence. Step 3: Provide the exact answer within the word limit.'
  };
  return strategies[type] || strategies.text;
}

function answersMatch_(answer, acceptedAnswers) {
  const normalize = value => String(value || '')
    .toLowerCase()
    .trim()
    .replace(/%/g, ' percent')
    .replace(/\bper\s+cent\b/g, ' percent')
    .replace(/\bpercent\b/g, ' percent')
    .replace(/\bninety\b/g, '90')
    .replace(/\bsix\b/g, '6')
    .replace(/[.,/#!$%^&*;:{}=-_`~()]/g, '')
    .replace(/\s+/g, ' ')
    .replace(/^the\s+/, '');

  const candidate = normalize(answer);
  return (acceptedAnswers || []).some(item => normalize(item) === candidate);
}

function saveFinalSummary_(email, testId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = ss.getSheetByName(PTE.SHEETS.RESPONSES);
  if (!responseSheet || responseSheet.getLastRow() < 2) {
    return {totalRaw: 0, scores: {listening: 10, reading: 10, writing: 10, speaking: 10}, grade: pteMigrationLevel_({listening: 10, reading: 10, writing: 10, speaking: 10})};
  }

  const rows = responseSheet.getDataRange().getValues();
  const scoreByBlock = {1: 0, 2: 0, 3: 0, 4: 0};
  const seen = {};

  for (let i = rows.length - 1; i >= 1; i--) {
    const row = rows[i];
    const uniqueKey = row[2] + '' + row[3] + '' + row[4];

    if (row[1] === email && row[2] === testId && !seen[uniqueKey]) {
      seen[uniqueKey] = true;
      if (row[6] === 'YES') scoreByBlock[Number(row[3])]++;
    }
  }

  const total = scoreByBlock[1] + scoreByBlock[2] + scoreByBlock[3] + scoreByBlock[4];

  const scores = {
    speaking: Math.round(10 + (scoreByBlock[1] / 15) * 80),
    writing: Math.round(10 + (scoreByBlock[2] / 5) * 80),
    reading: Math.round(10 + (scoreByBlock[3] / 15) * 80),
    listening: Math.round(10 + (scoreByBlock[4] / 15) * 80)
  };

  scores.speaking = Math.min(90, scores.speaking || 10);
  scores.writing = Math.min(90, scores.writing || 10);
  scores.reading = Math.min(90, scores.reading || 10);
  scores.listening = Math.min(90, scores.listening || 10);

  const grade = pteMigrationLevel_(scores);

  const resultSheet = ensureSheet_(ss, PTE.SHEETS.RESULTS, [ 'Timestamp', 'UserEmail', 'TestID', 'Block1', 'Block2', 'Block3', 'Block4', 'TotalOutOf50', 'BandScore', 'Band7Achieved' ]);

  resultSheet.appendRow([ new Date(), email, testId, scoreByBlock[1], scoreByBlock[2], scoreByBlock[3], scoreByBlock[4], total, grade.level, grade.met ? 'YES' : 'NO' ]);

  const bandCells = resultSheet.getRange(resultSheet.getLastRow(), 9, 1, 2);
  if (grade.met) bandCells.setBackground('#D1FAE5').setFontColor('#065F46').setFontWeight('bold');
  else bandCells.setBackground('#FEE2E2').setFontColor('#991B1B');

  return {totalRaw: total, scores: scores, grade: grade};
}

function pteMigrationLevel_(scores) {
  const value = key => Number((scores || {})[key] || 0);
  const listening = value('listening');
  const reading = value('reading');
  const writing = value('writing');
  const speaking = value('speaking');

  if (listening >= 69 && reading >= 70 && writing >= 85 && speaking >= 88) {
    return {level: 'Superior English', migrationPoints: 20, met: true, requirements: 'L69 · R70 · W85 · S88'};
  }
  if (listening >= 58 && reading >= 59 && writing >= 69 && speaking >= 76) {
    return {level: 'Proficient English', migrationPoints: 10, met: true, requirements: 'L58 · R59 · W69 · S76'};
  }
  if (listening >= 47 && reading >= 48 && writing >= 51 && speaking >= 54) {
    return {level: 'Competent English', migrationPoints: 0, met: true, requirements: 'L47 · R48 · W51 · S54'};
  }
  return {level: 'Below Competent English', migrationPoints: 0, met: false, requirements: 'L47 · R48 · W51 · S54'};
}

function getPteMigrationLevel(payload) {
  return pteMigrationLevel_((payload || {}).scores || {});
}

function getUserHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PTE.SHEETS.RESULTS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const email = Session.getActiveUser().getEmail() || 'anonymous_candidate';

  return sheet.getDataRange().getValues().slice(1)
    .filter(row => row[1] === email || email === 'anonymous_candidate')
    .reverse()
    .map(row => ({
      date: Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      testId: row[2], block1: row[3], block2: row[4], block3: row[5], block4: row[6], total: row[7], band: row[8], passed: row[9] === 'YES'
    }));
}

function getTestAnalysis(testId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responses = ss.getSheetByName(PTE.SHEETS.RESPONSES);
  if (!responses || responses.getLastRow() < 2) return { errors: [] };
  const rows = responses.getDataRange().getValues().slice(1);
  const email = Session.getActiveUser().getEmail() || 'anonymous_candidate';
  const errors = rows.filter(r => r[1] === email && r[2] === testId && r[6] === 'NO').map(r => ({
    qNum: r[4], userAnswer: r[5]
  }));
  return { errors: errors };
}


function generateSamplePteTest() {
  const testId = 'SAMPLE_' + Date.now();
  const passages = {
    p1Title: 'Speaking Tasks', p1Text: 'Read the prompts carefully and speak naturally.',
    p2Title: 'Writing Tasks', p2Text: 'Focus on spelling, grammar, and meeting the word counts.',
    p3Title: 'Reading Tasks', p3Text: 'Scan the text efficiently to find the correct information.',
    p4Title: 'Listening Tasks', p4Text: 'Listen to the audio or read the prompt and respond accurately.'
  };

  const sampleQuestions = [
    // Speaking
    { qNum: 1, blockId: 1, type: 'speaking', prompt: 'Read the following passage aloud: Climate change is one of the biggest challenges facing humanity. Governments, businesses, and individuals all have a role in reducing carbon emissions and protecting natural resources for future generations.', instruction: 'Read Aloud', acceptedAnswers: ['Practice estimate: 90 / 90'] },
    { qNum: 2, blockId: 1, type: 'speaking', prompt: 'Artificial intelligence is transforming industries by improving productivity and enabling smarter decision-making. However, ethical concerns remain an important topic for researchers.', instruction: 'Read Aloud', acceptedAnswers: ['Practice estimate: 90 / 90'] },
    { qNum: 3, blockId: 1, type: 'speaking', prompt: 'The conference has been postponed until next Friday.', instruction: 'Repeat Sentence', acceptedAnswers: ['Practice estimate: 90 / 90'] },
    { qNum: 4, blockId: 1, type: 'speaking', prompt: 'Students should submit their assignments before midnight.', instruction: 'Repeat Sentence', acceptedAnswers: ['Practice estimate: 90 / 90'] },
    { qNum: 5, blockId: 1, type: 'speaking', prompt: 'Public transportation helps reduce traffic congestion.', instruction: 'Repeat Sentence', acceptedAnswers: ['Practice estimate: 90 / 90'] },
    { qNum: 6, blockId: 1, type: 'speaking', prompt: 'Describe a bar chart showing: Australia – 35%, Canada – 25%, UK – 20%, USA – 15%, New Zealand – 5%', instruction: 'Describe Image', acceptedAnswers: ['Practice estimate: 90 / 90'] },
    { qNum: 7, blockId: 1, type: 'speaking', prompt: 'Describe a line graph showing smartphone users from 2018 to 2025 steadily increasing from 40 million to 90 million.', instruction: 'Describe Image', acceptedAnswers: ['Practice estimate: 90 / 90'] },
    { qNum: 8, blockId: 1, type: 'speaking', prompt: 'A lecture discusses renewable energy sources such as solar, wind, and hydroelectric power and explains their environmental benefits.', instruction: 'Re-tell Lecture', acceptedAnswers: ['Practice estimate: 90 / 90'] },
    { qNum: 9, blockId: 1, type: 'speaking', prompt: 'A professor explains the importance of effective time management for university students.', instruction: 'Re-tell Lecture', acceptedAnswers: ['Practice estimate: 90 / 90'] },
    { qNum: 10, blockId: 1, type: 'text', prompt: 'What do bees produce?', instruction: 'Answer Short Question', acceptedAnswers: ['Honey', 'honey'] },
    { qNum: 11, blockId: 1, type: 'text', prompt: 'Which planet is known as the Red Planet?', instruction: 'Answer Short Question', acceptedAnswers: ['Mars', 'mars'] },
    { qNum: 12, blockId: 1, type: 'text', prompt: 'What do you use to cut paper?', instruction: 'Answer Short Question', acceptedAnswers: ['Scissors', 'scissors'] },
    { qNum: 13, blockId: 1, type: 'text', prompt: 'Which animal is known as the king of the jungle?', instruction: 'Answer Short Question', acceptedAnswers: ['Lion', 'lion'] },
    { qNum: 14, blockId: 1, type: 'text', prompt: 'How many days are there in a leap year?', instruction: 'Answer Short Question', acceptedAnswers: ['366', '366 days'] },
    { qNum: 15, blockId: 1, type: 'text', prompt: 'What is the opposite of "ancient"?', instruction: 'Answer Short Question', acceptedAnswers: ['Modern', 'modern', 'new'] },

    // Writing
    { qNum: 16, blockId: 2, type: 'writing', prompt: 'Read a 250-word passage about online education and summarize it in one sentence (5–75 words).', instruction: 'Summarize Written Text', acceptedAnswers: ['sample summary'] },
    { qNum: 17, blockId: 2, type: 'writing', prompt: 'Read a passage discussing climate change and summarize it in one sentence.', instruction: 'Summarize Written Text', acceptedAnswers: ['sample summary'] },
    { qNum: 18, blockId: 2, type: 'writing', prompt: 'Should university education be free for all students? Write 200–300 words.', instruction: 'Essay Topic', acceptedAnswers: ['sample essay'] },
    { qNum: 19, blockId: 2, type: 'writing', prompt: 'Do the advantages of working from home outweigh the disadvantages?', instruction: 'Essay Topic', acceptedAnswers: ['sample essay'] },
    { qNum: 20, blockId: 2, type: 'writing', prompt: 'Some people believe technology makes life easier, while others think it creates problems. Discuss both views.', instruction: 'Essay Topic', acceptedAnswers: ['sample essay'] },

    // Reading
    { qNum: 21, blockId: 3, type: 'multiple_choice', prompt: 'Which gas do plants absorb?', instruction: 'Multiple Choice (Single Answer)', options: ['Oxygen', 'Carbon Dioxide', 'Nitrogen', 'Hydrogen'], acceptedAnswers: ['Carbon Dioxide'] },
    { qNum: 22, blockId: 3, type: 'multiple_choice', prompt: 'Which are renewable energy sources?', instruction: 'Multiple Choice (Multiple Answers)', options: ['Solar', 'Wind', 'Coal', 'Hydroelectric'], acceptedAnswers: ['Solar, Wind, Hydroelectric'] },
    { qNum: 23, blockId: 3, type: 'text', prompt: 'Arrange: Finally, the report was published. | Researchers collected data. | The project began in January. | The findings were analyzed.', instruction: 'Re-order Paragraph (Type sequence)', acceptedAnswers: ['The project began in January. Researchers collected data. The findings were analyzed. Finally, the report was published.'] },
    { qNum: 24, blockId: 3, type: 'multiple_choice', prompt: 'Exercise regularly because it helps improve your .', instruction: 'Fill in the Blanks', options: ['Health', 'Weather', 'Traffic', 'Furniture'], acceptedAnswers: ['Health'] },
    { qNum: 25, blockId: 3, type: 'multiple_choice', prompt: 'The scientist carefully the experiment.', instruction: 'Fill in the Blanks', options: ['Conducted', 'Ate', 'Bought', 'Slept'], acceptedAnswers: ['Conducted'] },
    { qNum: 26, blockId: 3, type: 'multiple_choice', prompt: 'The company plans to production next year.', instruction: 'Reading & Writing Fill in the Blanks', options: ['Increase', 'Break', 'Sleep', 'Wash'], acceptedAnswers: ['Increase'] },
    { qNum: 27, blockId: 3, type: 'multiple_choice', prompt: 'The museum attracts thousands of every year.', instruction: 'Reading & Writing Fill in the Blanks', options: ['Visitors', 'Doctors', 'Cars', 'Rivers'], acceptedAnswers: ['Visitors'] },
    { qNum: 28, blockId: 3, type: 'multiple_choice', prompt: 'Why is recycling important?', instruction: 'Multiple Choice', options: ['Saves resources', 'Increases pollution', 'Wastes energy', 'Reduces education'], acceptedAnswers: ['Saves resources'] },
    { qNum: 29, blockId: 3, type: 'text', prompt: 'Arrange: She completed her research. | She enrolled in university. | She published her paper. | She graduated.', instruction: 'Re-order Paragraph (Type sequence)', acceptedAnswers: ['She enrolled in university. She completed her research. She graduated. She published her paper.'] },
    { qNum: 30, blockId: 3, type: 'multiple_choice', prompt: 'The meeting was due to bad weather.', instruction: 'Fill in the Blank', options: ['Postponed', 'Cooked', 'Painted', 'Jumped'], acceptedAnswers: ['Postponed'] },
    { qNum: 31, blockId: 3, type: 'multiple_choice', prompt: "The manager appreciated the employee's .", instruction: 'Fill in the Blank', options: ['Dedication', 'Rain', 'Window', 'Bicycle'], acceptedAnswers: ['Dedication'] },
    { qNum: 32, blockId: 3, type: 'multiple_choice', prompt: 'Education provides people with valuable .', instruction: 'Reading Fill in Blank', options: ['Knowledge', 'Chairs', 'Clouds', 'Cars'], acceptedAnswers: ['Knowledge'] },
    { qNum: 33, blockId: 3, type: 'multiple_choice', prompt: 'Which animal is a mammal?', instruction: 'Multiple Choice', options: ['Dolphin', 'Shark', 'Trout', 'Octopus'], acceptedAnswers: ['Dolphin'] },
    { qNum: 34, blockId: 3, type: 'multiple_choice', prompt: 'Scientists continue to new medicines.', instruction: 'Reading Fill in Blank', options: ['Develop', 'Forget', 'Destroy', 'Ignore'], acceptedAnswers: ['Develop'] },
    { qNum: 35, blockId: 3, type: 'text', prompt: 'Arrange: Seeds were planted. | Plants started growing. | Flowers bloomed. | Fruits appeared.', instruction: 'Re-order Paragraph (Type sequence)', acceptedAnswers: ['Seeds were planted. Plants started growing. Flowers bloomed. Fruits appeared.'] },

    // Listening
    { qNum: 36, blockId: 4, type: 'writing', prompt: 'Listen to a lecture about global warming and summarize it in 50–70 words.', instruction: 'Summarize Spoken Text', acceptedAnswers: ['sample summary'] },
    { qNum: 37, blockId: 4, type: 'multiple_choice', prompt: 'The speaker mainly discusses:', instruction: 'Multiple Choice', options: ['Education', 'Climate Change', 'Sports', 'Music'], acceptedAnswers: ['Climate Change'] },
    { qNum: 38, blockId: 4, type: 'text', prompt: 'The professor explained that renewable energy reduces emissions.', instruction: 'Fill in the Blanks (Type missing words)', acceptedAnswers: ['carbon', 'greenhouse gas'] },
    { qNum: 39, blockId: 4, type: 'multiple_choice', prompt: 'Choose the summary that best matches the lecture.', instruction: 'Highlight Correct Summary', options: ['Correct Summary', 'Incorrect 1', 'Incorrect 2', 'Incorrect 3'], acceptedAnswers: ['Correct Summary'] },
    { qNum: 40, blockId: 4, type: 'multiple_choice', prompt: 'Select the benefits of exercise mentioned by the speaker.', instruction: 'Multiple Choice (Multiple Answers)', options: ['Health', 'Wealth', 'Fitness', 'Fame'], acceptedAnswers: ['Health', 'Fitness'] },
    { qNum: 41, blockId: 4, type: 'multiple_choice', prompt: 'The lecture ends with: "The future of education depends on continuous ."', instruction: 'Select Missing Word', options: ['learning', 'sleeping', 'eating', 'driving'], acceptedAnswers: ['learning'] },
    { qNum: 42, blockId: 4, type: 'text', prompt: 'Listen and identify words that differ from the transcript.', instruction: 'Highlight Incorrect Words', acceptedAnswers: ['incorrect'] },
    { qNum: 43, blockId: 4, type: 'text', prompt: 'Education is the foundation of a successful career.', instruction: 'Write From Dictation', acceptedAnswers: ['Education is the foundation of a successful career.'] },
    { qNum: 44, blockId: 4, type: 'text', prompt: 'Technology has changed the way people communicate.', instruction: 'Write From Dictation', acceptedAnswers: ['Technology has changed the way people communicate.'] },
    { qNum: 45, blockId: 4, type: 'text', prompt: 'Regular exercise improves both physical and mental health.', instruction: 'Write From Dictation', acceptedAnswers: ['Regular exercise improves both physical and mental health.'] },
    { qNum: 46, blockId: 4, type: 'text', prompt: 'Climate change affects every country in the world.', instruction: 'Write From Dictation', acceptedAnswers: ['Climate change affects every country in the world.'] },
    { qNum: 47, blockId: 4, type: 'text', prompt: 'Students should practice English every day.', instruction: 'Write From Dictation', acceptedAnswers: ['Students should practice English every day.'] },
    { qNum: 48, blockId: 4, type: 'text', prompt: 'Innovation drives economic growth and development.', instruction: 'Write From Dictation', acceptedAnswers: ['Innovation drives economic growth and development.'] },
    { qNum: 49, blockId: 4, type: 'text', prompt: 'Reading books helps improve vocabulary and comprehension.', instruction: 'Write From Dictation', acceptedAnswers: ['Reading books helps improve vocabulary and comprehension.'] },
    { qNum: 50, blockId: 4, type: 'text', prompt: 'Success requires patience, dedication, and hard work.', instruction: 'Write From Dictation', acceptedAnswers: ['Success requires patience, dedication, and hard work.'] }
  ];

  const draft = {
    testId: testId,
    title: 'PTE 50-Question Sample Exam (2026)',
    sourceFileName: 'Internal Generation',
    status: 'DRAFT — READY FOR REVIEW',
    passages: passages,
    settings: {
      requireReadingCompletion: true,
      immediateFeedback: true,
      showRunningScore: true,
      enablePostBlockReview: true,
      published: false
    },
    questions: normaliseQuestions_(sampleQuestions)
  };

  saveNewDraft_(draft);
  return getAdminDraft(draft.testId);
}

function getOverallAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheet = ss.getSheetByName(PTE.SHEETS.RESULTS);
  if (!resultSheet || resultSheet.getLastRow() < 2) return {status: 'error', message: 'No test history available.'};

  const email = Session.getActiveUser().getEmail() || 'anonymous_candidate';
  const rows = resultSheet.getDataRange().getValues().slice(1).filter(r => r[1] === email || email === 'anonymous_candidate');

  if (rows.length === 0) return {status: 'error', message: 'No test history available for this user.'};

  let totals = {speaking: 0, writing: 0, reading: 0, listening: 0, count: 0};

  rows.forEach(r => {
    // block 1 is speaking, 2 is writing, 3 is reading, 4 is listening
    totals.speaking += Number(r[3] || 0);
    totals.writing += Number(r[4] || 0);
    totals.reading += Number(r[5] || 0);
    totals.listening += Number(r[6] || 0);
    totals.count++;
  });

  const avg = {
    speaking: Math.round(10 + ((totals.speaking / totals.count) / 15) * 80),
    writing: Math.round(10 + ((totals.writing / totals.count) / 5) * 80),
    reading: Math.round(10 + ((totals.reading / totals.count) / 15) * 80),
    listening: Math.round(10 + ((totals.listening / totals.count) / 15) * 80)
  };

  const grade = pteMigrationLevel_(avg);

  // Provide targeted feedback based on the weakest area
  let weakest = Object.keys(avg).reduce((a, b) => avg[a] < avg[b] ? a : b);
  let feedback = '';

  if (weakest === 'speaking') {
    feedback = "Your speaking score is the lowest. Focus on oral fluency and pronunciation. Practice reading aloud daily, and do not hesitate or self-correct during the exam.";
  } else if (weakest === 'writing') {
    feedback = "Your writing score needs improvement. Focus on adhering strictly to word limits, grammatical accuracy, and spelling. Review the structure for summarizing written text and writing essays.";
  } else if (weakest === 'reading') {
    feedback = "Your reading score is the weakest area. Practice scanning and skimming techniques. Work on understanding paragraph structures for Re-order Paragraphs and multiple-choice questions.";
  } else if (weakest === 'listening') {
    feedback = "Your listening score requires attention. Practice active listening and note-taking. Pay special attention to Write From Dictation and Summarize Spoken Text.";
  }

  return {
    averages: avg,
    grade: grade,
    weakestArea: weakest,
    constructiveFeedback: feedback
  };
}

function generatePackageFromPdf_(pdfBase64, fileName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY is required for PDF processing.');

  const model = PropertiesService.getScriptProperties().getProperty('GEMINI_VISUAL_MODEL') || GEMINI_VISUAL_MODEL;

  const prompt = `You are a PTE Exam generation system. Extract the test content from the provided PDF into a structured JSON package.
  Ensure the JSON matches this exact structure:
  {
    "title": "Extracted PTE Exam",
    "blocks": [
      {
        "blockId": 1,
        "passageTitle": "Speaking Tasks",
        "passageText": "...",
        "questions": [
           {"qNum": 1, "type": "speaking", "prompt": "...", "instruction": "...", "acceptedAnswers": ["Practice estimate: 90 / 90"]}
        ]
      }
    ]
  }
  Create exactly 50 questions distributed across 4 blocks (1: Speaking Q1-15, 2: Writing Q16-20, 3: Reading Q21-35, 4: Listening Q36-50).
  Ensure all questions have prompt, instruction, and acceptedAnswers. Multiple choice must have an options array.
  Do not include markdown blocks, just return raw JSON.`;

  const request = {
    contents: [{parts: [
      {inline_data: {mime_type: 'application/pdf', data: String(pdfBase64)}},
      {text: prompt}
    ]}],
    generationConfig: {responseMimeType: 'application/json', temperature: 0.1}
  };

  const response = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(apiKey),
    {method: 'post', contentType: 'application/json', payload: JSON.stringify(request), muteHttpExceptions: true}
  );

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error('Gemini API failed: ' + response.getContentText().slice(0, 500));
  }

  let imported;
  try {
    const body = JSON.parse(response.getContentText());
    const candidate = body.candidates && body.candidates[0];
    const parts = candidate && candidate.content && candidate.content.parts;
    const text = (parts || []).map(part => part.text || '').join('');
    imported = JSON.parse(text);
  } catch (err) {
    throw new Error('Failed to parse Gemini JSON response: ' + err.message);
  }

  return buildDraftFromPackage_(imported, fileName);
}
