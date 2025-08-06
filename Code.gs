/**
 * 集合知AIサービス - Google Apps Script メインファイル
 * 会議やセミナーで参加者がAIと対話し、意見を収集・集約するシステム
 */

// シート名定数
const SHEET_NAMES = {
  QUESTION_TEMPLATES: 'QuestionTemplates',
  QUESTIONS: 'Questions', 
  RESULTS: 'Results',
  SUMMARY: 'Summary'
};

// デバッグ機能
const DEBUG = {
  enabled: () => {
    const debugMode = PropertiesService.getScriptProperties().getProperty('DEBUG_MODE');
    return debugMode === 'true';
  },
  
  log: (message, data = null) => {
    const timestamp = new Date().toLocaleString('ja-JP');
    const logMessage = `[${timestamp}] ${message}`;
    
    console.log(logMessage);
    if (data) {
      console.log('データ:', JSON.stringify(data, null, 2));
    }
    
    // デバッグモードが有効な場合、より詳細なログを出力
    if (DEBUG.enabled()) {
      const caller = DEBUG.getCaller();
      console.log(`[DEBUG] 呼び出し元: ${caller}`);
      
      // スタックトレース
      try {
        throw new Error();
      } catch (e) {
        const stack = e.stack.split('\n').slice(2, 5).join('\n  ');
        console.log(`[DEBUG] スタック:\n  ${stack}`);
      }
    }
  },
  
  error: (message, error = null) => {
    const timestamp = new Date().toLocaleString('ja-JP');
    const errorMessage = `[ERROR ${timestamp}] ${message}`;
    
    console.error(errorMessage);
    if (error) {
      console.error('エラー詳細:', error.toString());
      console.error('スタックトレース:', error.stack);
    }
    
    // 重要なエラーはスプレッドシートにも記録
    if (DEBUG.enabled()) {
      DEBUG.logToSheet('ERROR', message, error);
    }
  },
  
  warn: (message, data = null) => {
    const timestamp = new Date().toLocaleString('ja-JP');
    console.warn(`[WARN ${timestamp}] ${message}`);
    if (data) {
      console.warn('データ:', data);
    }
  },
  
  getCaller: () => {
    try {
      throw new Error();
    } catch (e) {
      const stack = e.stack.split('\n');
      const callerLine = stack[4] || stack[3] || 'unknown';
      return callerLine.trim();
    }
  },
  
  logToSheet: (level, message, data = null) => {
    try {
      const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      if (!spreadsheetId) return;
      
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      let logSheet;
      
      try {
        logSheet = spreadsheet.getSheetByName('DebugLog');
      } catch (e) {
        logSheet = spreadsheet.insertSheet('DebugLog');
        logSheet.getRange('A1:E1').setValues([
          ['Timestamp', 'Level', 'Message', 'Data', 'Caller']
        ]);
      }
      
      logSheet.appendRow([
        new Date(),
        level,
        message,
        data ? JSON.stringify(data) : '',
        DEBUG.getCaller()
      ]);
    } catch (e) {
      console.error('ログシートへの書き込みエラー:', e);
    }
  },
  
  measure: (label, func) => {
    const start = new Date();
    DEBUG.log(`[MEASURE] ${label} 開始`);
    
    try {
      const result = func();
      const duration = new Date() - start;
      DEBUG.log(`[MEASURE] ${label} 完了 (${duration}ms)`);
      return result;
    } catch (error) {
      const duration = new Date() - start;
      DEBUG.error(`[MEASURE] ${label} エラー (${duration}ms)`, error);
      throw error;
    }
  }
};

/**
 * スプレッドシート構造の初期セットアップ
 * 新しいスプレッドシートに必要なシートとヘッダーを作成
 */
function setupSpreadsheetStructure() {
  DEBUG.log('スプレッドシート構造セットアップ開始');
  
  // スプレッドシートIDを取得
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  DEBUG.log('スプレッドシートID取得', { spreadsheetId: spreadsheetId ? spreadsheetId.substring(0, 10) + '...' : 'null' });
  
  if (!spreadsheetId) {
    DEBUG.error('SPREADSHEET_IDが設定されていません');
    throw new Error('SPREADSHEET_IDが設定されていません。先にスクリプトプロパティを設定してください。');
  }
  
  let spreadsheet;
  try {
    DEBUG.log('スプレッドシートアクセス試行');
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    DEBUG.log('スプレッドシートアクセス成功', { name: spreadsheet.getName() });
  } catch (error) {
    DEBUG.error('スプレッドシートアクセスエラー', error);
    throw new Error(`スプレッドシートにアクセスできません (ID: ${spreadsheetId}): ${error.message}`);
  }
  
  try {
    // 必要なシートがすでに存在するかチェック
    DEBUG.log('既存シートの確認開始');
    const existingSheets = spreadsheet.getSheets();
    const existingSheetNames = existingSheets.map(sheet => sheet.getName());
    const requiredSheetNames = Object.values(SHEET_NAMES);
    
    DEBUG.log('シート状況確認', { 
      existing: existingSheetNames, 
      required: requiredSheetNames 
    });
    
    // 必要なシートが既に全て存在する場合は処理をスキップ
    const missingSheets = requiredSheetNames.filter(name => !existingSheetNames.includes(name));
    DEBUG.log('不足シート確認', { missing: missingSheets });
    
    if (missingSheets.length === 0) {
      DEBUG.log('必要なシートは既に存在 - スキップ');
      return { success: true, message: '必要なシートは既に存在しています' };
    }
    
    const createdSheets = [];
    DEBUG.log('不足シートの作成開始', { count: missingSheets.length });
    
    // 必要なシートのみを作成
    if (missingSheets.includes(SHEET_NAMES.QUESTION_TEMPLATES)) {
      DEBUG.log('QuestionTemplatesシート作成中');
      const templateSheet = spreadsheet.insertSheet(SHEET_NAMES.QUESTION_TEMPLATES);
      templateSheet.getRange('A1:F1').setValues([
        ['template_id', 'template_name', 'theme', 'question_type', 'question_text', 'created_at']
      ]);
      createdSheets.push(templateSheet);
      DEBUG.log('QuestionTemplatesシート作成完了');
    }
    
    if (missingSheets.includes(SHEET_NAMES.QUESTIONS)) {
      const questionSheet = spreadsheet.insertSheet(SHEET_NAMES.QUESTIONS);
      questionSheet.getRange('A1:F1').setValues([
        ['session_id', 'theme', 'question_type', 'question_text', 'source_type', 'created_at']
      ]);
      createdSheets.push(questionSheet);
    }
    
    if (missingSheets.includes(SHEET_NAMES.RESULTS)) {
      const resultSheet = spreadsheet.insertSheet(SHEET_NAMES.RESULTS);
      resultSheet.getRange('A1:G1').setValues([
        ['session_id', 'participant_id', 'question_id', 'question_type', 'user_input', 'ai_response', 'timestamp']
      ]);
      createdSheets.push(resultSheet);
    }
    
    if (missingSheets.includes(SHEET_NAMES.SUMMARY)) {
      const summarySheet = spreadsheet.insertSheet(SHEET_NAMES.SUMMARY);
      summarySheet.getRange('A1:G1').setValues([
        ['session_id', 'theme', 'participant_count', 'consensus_points', 'divergent_points', 'key_insights', 'full_summary']
      ]);
      createdSheets.push(summarySheet);
    }
    
    // 作成されたシートのヘッダー行スタイル設定
    createdSheets.forEach(sheet => {
      const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e8f0fe');
      headerRange.setBorder(true, true, true, true, true, true);
    });
    
    const message = createdSheets.length > 0 ? 
      `${createdSheets.length}個のシートを作成しました` :
      '必要なシートは既に存在していました';
      
    return { success: true, message: message, createdSheets: createdSheets.length };
    
  } catch (error) {
    return { success: false, message: 'スプレッドシート設定中にエラーが発生しました: ' + error.toString() };
  }
}

/**
 * 指定されたシートを取得するヘルパー関数
 * @param {string} sheetName - シート名
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} シートオブジェクト
 */
function getSheet(sheetName) {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  
  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_IDが設定されていません。初期設定を完了してください。');
  }
  
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`シート "${sheetName}" が見つかりません`);
    }
    
    return sheet;
  } catch (error) {
    console.error(`シート取得エラー (${sheetName}):`, error);
    throw error;
  }
}

/**
 * デフォルト質問テンプレートの初期化
 * システム利用開始時に実行
 */
function initializeDefaultTemplates() {
  try {
    // 新入社員研修用テンプレート
    saveQuestionTemplate(
      '新入社員研修用',
      '新入社員の意識・期待調査', 
      [
        '当社に入社した理由や動機を教えてください',
        '新入社員研修で学びたいことや期待することは何ですか？',
        '理想の社会人像や、どのような成長を目指したいか話し合いましょう'
      ]
    );
    
    // 製品企画会議用テンプレート
    saveQuestionTemplate(
      '製品企画会議用',
      '新製品アイデア・要望収集',
      [
        '現在の製品について、改善すべき点はありますか？',
        '顧客から受けた要望や意見で印象に残っているものを教えてください',
        '理想的な新製品について、自由にアイデアを話し合いましょう'
      ]
    );
    
    // 業務改善用テンプレート
    saveQuestionTemplate(
      '業務改善ディスカッション用',
      '業務効率化・改善提案',
      [
        '現在の業務で時間がかかりすぎていると感じる作業はありますか？',
        '他部署との連携で困っていることや改善したい点はありますか？',
        '理想的な職場環境や働き方について話し合いましょう'
      ]
    );
    
    return { success: true, message: 'デフォルトテンプレートを正常に作成しました' };
    
  } catch (error) {
    return { success: false, message: 'テンプレート初期化中にエラーが発生しました: ' + error.toString() };
  }
}

/**
 * 質問テンプレートを保存
 * @param {string} templateName - テンプレート名
 * @param {string} theme - テーマ
 * @param {Array<string>} questions - 質問配列（3つ）
 * @returns {string} テンプレートID
 */
function saveQuestionTemplate(templateName, theme, questions) {
  if (!questions || questions.length !== 3) {
    throw new Error('質問は3つ（定型2問 + 深掘り1問）必要です');
  }
  
  const templateId = Utilities.getUuid();
  const sheet = getSheet(SHEET_NAMES.QUESTION_TEMPLATES);
  const questionTypes = ['fixed_1', 'fixed_2', 'free_discussion'];
  const now = new Date();
  
  // 各質問を1行ずつ追加
  questions.forEach((questionText, index) => {
    sheet.appendRow([
      templateId,
      templateName,  
      theme,
      questionTypes[index],
      questionText,
      now
    ]);
  });
  
  return templateId;
}

/**
 * 初期設定の完全実行
 * スプレッドシート構造作成 + デフォルトテンプレート初期化
 */
function completeInitialSetup() {
  try {
    // ステップ1: スプレッドシート構造作成
    const structureResult = setupSpreadsheetStructure();
    if (!structureResult.success) {
      throw new Error('スプレッドシート構造作成に失敗: ' + structureResult.message);
    }
    
    // ステップ2: デフォルトテンプレート初期化
    const templateResult = initializeDefaultTemplates();
    if (!templateResult.success) {
      throw new Error('デフォルトテンプレート作成に失敗: ' + templateResult.message);
    }
    
    return {
      success: true,
      message: 'システムの初期設定が正常に完了しました',
      structureResult: structureResult,
      templateResult: templateResult
    };
    
  } catch (error) {
    return {
      success: false,
      message: '初期設定中にエラーが発生しました: ' + error.toString()
    };
  }
}

/**
 * システム状態確認
 * 設定が正しく完了しているかチェック
 */
function checkSystemStatus() {
  const status = {
    spreadsheetStructure: false,
    apiKeyConfigured: false,
    defaultTemplatesLoaded: false,
    errors: []
  };
  
  try {
    // スプレッドシート構造確認
    const sheetNames = Object.values(SHEET_NAMES);
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    sheetNames.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        status.errors.push(`シート "${sheetName}" が見つかりません`);
      }
    });
    
    if (status.errors.length === 0) {
      status.spreadsheetStructure = true;
    }
    
    // API Key設定確認
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (apiKey && apiKey.length > 10) {
      status.apiKeyConfigured = true;
    } else {
      status.errors.push('Gemini API Keyが設定されていません');
    }
    
    // デフォルトテンプレート確認
    const templateSheet = spreadsheet.getSheetByName(SHEET_NAMES.QUESTION_TEMPLATES);
    if (templateSheet && templateSheet.getLastRow() > 1) {
      status.defaultTemplatesLoaded = true;
    } else {
      status.errors.push('デフォルトテンプレートが読み込まれていません');
    }
    
  } catch (error) {
    status.errors.push('システム状態チェック中にエラー: ' + error.toString());
  }
  
  return status;
}

// ============================================================================
// セッション管理機能
// ============================================================================

/**
 * セッション作成（3つのモード対応）
 * @param {string} theme - セッションテーマ
 * @param {string} questionMode - 質問設定モード ('ai_generated', 'template', 'custom')
 * @param {string} templateId - テンプレートID（templateモードの場合）
 * @param {Array<string>} customQuestions - カスタム質問配列（customモードの場合）
 * @returns {Object} 結果オブジェクト
 */
function createSession(theme, questionMode = 'ai_generated', templateId = null, customQuestions = null) {
  DEBUG.log('セッション作成開始', { 
    theme, 
    questionMode, 
    templateId: templateId ? templateId.substring(0, 8) + '...' : null,
    hasCustomQuestions: !!customQuestions 
  });
  
  const sessionId = Utilities.getUuid();
  DEBUG.log('セッションID生成', { sessionId });
  
  let questions;
  let sourceType;
  
  try {
    DEBUG.log('質問生成開始', { questionMode });
    
    switch (questionMode) {
      case 'ai_generated':
        DEBUG.log('AI質問生成モード');
        if (customQuestions && customQuestions.length >= 3) {
          DEBUG.log('カスタム質問を使用（AI生成済み）', { count: customQuestions.length });
          questions = customQuestions;
        } else {
          DEBUG.log('サーバーサイドでAI質問生成');
          questions = generateQuestionsWithGemini(theme);
        }
        sourceType = 'ai_generated';
        DEBUG.log('AI質問設定完了', { questions });
        break;
        
      case 'template':
        if (!templateId) {
          DEBUG.error('テンプレートIDが未指定');
          throw new Error('テンプレートIDが指定されていません');
        }
        DEBUG.log('テンプレート読み込みモード', { templateId });
        questions = loadQuestionsFromTemplate(templateId);
        sourceType = 'template';
        DEBUG.log('テンプレート読み込み完了', { questions });
        break;
        
      case 'custom':
        if (!customQuestions || customQuestions.length !== 3) {
          DEBUG.error('カスタム質問が無効', { customQuestions });
          throw new Error('カスタム質問は3つ必要です');
        }
        DEBUG.log('カスタム質問モード');
        questions = customQuestions;
        sourceType = 'custom_edited';
        DEBUG.log('カスタム質問設定完了', { questions });
        break;
        
      default:
        DEBUG.error('無効な質問モード', { questionMode });
        throw new Error('無効な質問モードです: ' + questionMode);
    }
    
    // Questionsシートに保存
    const sheet = getSheet(SHEET_NAMES.QUESTIONS);
    const questionTypes = ['fixed_1', 'fixed_2', 'free_discussion'];
    const now = new Date();
    
    // 質問を配列形式で正規化
    const questionArray = Array.isArray(questions) ? questions : [
      questions.fixed_1 || questions[0],
      questions.fixed_2 || questions[1], 
      questions.free_discussion || questions[2]
    ];
    
    questionArray.forEach((questionText, index) => {
      if (questionText) {
        sheet.appendRow([
          sessionId, 
          theme, 
          questionTypes[index], 
          questionText, 
          sourceType,
          now
        ]);
      }
    });
    
    return {
      success: true,
      sessionId: sessionId,
      questions: questionArray,
      sourceType: sourceType,
      message: 'セッションを正常に作成しました'
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'セッション作成エラー: ' + error.toString()
    };
  }
}

/**
 * テンプレート一覧取得
 * @returns {Array<Object>} テンプレート一覧
 */
function getQuestionTemplates() {
  try {
    const sheet = getSheet(SHEET_NAMES.QUESTION_TEMPLATES);
    const data = sheet.getDataRange().getValues();
    const templates = new Map();
    
    // ヘッダー行をスキップ
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const templateId = row[0];
      const templateName = row[1];
      const theme = row[2];
      
      if (!templates.has(templateId)) {
        templates.set(templateId, {
          id: templateId,
          name: templateName,
          theme: theme,
          createdAt: row[5]
        });
      }
    }
    
    return Array.from(templates.values());
  } catch (error) {
    throw new Error('テンプレート一覧取得エラー: ' + error.toString());
  }
}

/**
 * テンプレート読み込み機能
 * @param {string} templateId - テンプレートID
 * @returns {Array<string>} 質問配列
 */
function loadQuestionsFromTemplate(templateId) {
  try {
    const sheet = getSheet(SHEET_NAMES.QUESTION_TEMPLATES);
    const data = sheet.getDataRange().getValues();
    const questions = [];
    
    // question_typeの順序で並び替え
    const questionOrder = ['fixed_1', 'fixed_2', 'free_discussion'];
    const orderedQuestions = new Array(3);
    
    // ヘッダー行をスキップして検索
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === templateId) {
        const typeIndex = questionOrder.indexOf(row[3]); // question_type列
        if (typeIndex !== -1) {
          orderedQuestions[typeIndex] = row[4]; // question_text列
        }
      }
    }
    
    return orderedQuestions;
  } catch (error) {
    throw new Error('テンプレート読み込みエラー: ' + error.toString());
  }
}

/**
 * 質問生成機能（Gemini API使用）
 * @param {string} theme - テーマ
 * @returns {Object} 生成された質問オブジェクト
 */
function generateQuestionsWithGemini(theme) {
  DEBUG.log('Gemini API質問生成開始', { theme });
  
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const endpoint = PropertiesService.getScriptProperties().getProperty('GEMINI_API_ENDPOINT') || 
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent';
  
  DEBUG.log('API設定確認', { 
    hasApiKey: !!apiKey,
    apiKeyLength: apiKey ? apiKey.length : 0,
    endpoint 
  });
  
  if (!apiKey) {
    DEBUG.error('Gemini API Keyが未設定');
    throw new Error('Gemini API Keyが設定されていません');
  }
  
  const prompt = `
テーマ「${theme}」について、会議やセミナーで参加者の意見を効率的に収集するための質問を生成してください。

以下の形式で3つの質問を作成してください：
1. 定型質問1: 基本的な立場・意見を聞く質問
2. 定型質問2: 具体的な経験・事例を聞く質問  
3. 深掘り質問: AIとの対話で議論を深められる開放的な質問

質問は80文字程度にしてください
定型質問1と定型質問2は例を示してください

JSON形式で出力してください：
{
  "fixed_1": "質問文",
  "fixed_2": "質問文", 
  "free_discussion": "質問文"
}
`;

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.8,
      maxOutputTokens: 1500,
    }
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-goog-api-key': apiKey
    },
    payload: JSON.stringify(payload)
  };
  
  try {
    DEBUG.log('Gemini API呼び出し開始');
    const response = UrlFetchApp.fetch(endpoint, options);
    const statusCode = response.getResponseCode();
    
    DEBUG.log('Gemini API応答受信', { 
      statusCode, 
      contentLength: response.getContentText().length 
    });
    
    if (statusCode !== 200) {
      DEBUG.error('Gemini API エラー応答', { 
        statusCode, 
        response: response.getContentText() 
      });
      throw new Error(`API Error: ${statusCode}`);
    }
    
    const data = JSON.parse(response.getContentText());
    const generatedText = data.candidates[0].content.parts[0].text;
    
    DEBUG.log('生成テキスト取得', { 
      textLength: generatedText.length,
        textPreview: generatedText.substring(0, 200),
        fullText: generatedText
      });
      
      // JSONを抽出・パース（より柔軟な抽出）
      let jsonMatch = generatedText.match(/\{[\s\S]*\}/);
      
      if (!jsonMatch) {
        // バックアップ: ```json から ``` までを抽出
        const codeBlockMatch = generatedText.match(/```json\s*([\s\S]*?)\s*```/);
        if (codeBlockMatch) {
          DEBUG.log('コードブロック形式からJSON抽出', { extractedJson: codeBlockMatch[1] });
          try {
            const parsedQuestions = JSON.parse(codeBlockMatch[1]);
            DEBUG.log('質問JSON解析成功（コードブロック）', parsedQuestions);
            return parsedQuestions;
          } catch (parseError) {
            DEBUG.error('コードブロックJSON解析エラー', parseError);
          }
        }
        
        // さらなるバックアップ: { から最後までを取得
        const partialJsonMatch = generatedText.match(/\{[\s\S]*/);
        if (partialJsonMatch) {
          DEBUG.log('部分的JSONを試行', { partialJson: partialJsonMatch[0] });
          try {
            // 不完全なJSONを補完して試行
            let jsonStr = partialJsonMatch[0];
            // 基本的な補完を試行
            if (!jsonStr.includes('"free_discussion"')) {
              jsonStr += '\n"free_discussion": "このテーマについて自由に議論しましょう。"\n}';
            }
            if (!jsonStr.endsWith('}')) {
              jsonStr += '}';
            }
            
            const parsedQuestions = JSON.parse(jsonStr);
            DEBUG.log('質問JSON解析成功（補完）', parsedQuestions);
            return parsedQuestions;
          } catch (parseError) {
            DEBUG.error('補完JSON解析エラー', parseError);
          }
        }
      } else {
        try {
      const parsedQuestions = JSON.parse(jsonMatch[0]);
      DEBUG.log('質問JSON解析成功', parsedQuestions);
      return parsedQuestions;
        } catch (parseError) {
          DEBUG.error('標準JSON解析エラー', parseError);
        }
    }
    
      DEBUG.warn('JSON形式での回答が得られませんでした', { 
        generatedText,
        textLength: generatedText.length 
      });
    throw new Error('JSON形式での回答が得られませんでした');
    
  } catch (error) {
    DEBUG.error('Gemini API呼び出しエラー', error);
    DEBUG.log('フォールバック質問を使用');
    
    // フォールバック用のデフォルト質問
    return {
      "fixed_1": `「${theme}」について、あなたの基本的な考えや立場を教えてください。`,
      "fixed_2": `「${theme}」に関連する具体的な経験や事例があれば共有してください。`,  
      "free_discussion": `「${theme}」についてもう少し深く議論してみましょう。どの観点から話したいですか？`
    };
  }
}

// ============================================================================
// WebApp エンドポイント
// ============================================================================

/**
 * デプロイIDを取得する関数
 * @returns {string} デプロイID
 */
function getDeploymentId() {
  try {
    // 方法1: PropertiesServiceから手動設定されたデプロイIDを取得
    const deploymentId = PropertiesService.getScriptProperties().getProperty('DEPLOYMENT_ID');
    if (deploymentId) {
      DEBUG.log('PropertiesServiceからデプロイID取得成功', { deploymentId: deploymentId.substring(0, 20) + '...' });
      return deploymentId;
    }
    
    // 方法2: ScriptApp.getService().getUrl()を試す（問題がある場合があるが試す価値はある）
    try {
      const webAppUrl = ScriptApp.getService().getUrl();
      if (webAppUrl) {
        const match = webAppUrl.match(/\/macros\/s\/([^\/]+)\//);
        if (match && match[1]) {
          const extractedId = match[1];
          DEBUG.log('ScriptApp.getService().getUrl()からデプロイID抽出成功', { deploymentId: extractedId.substring(0, 20) + '...' });
          return extractedId;
        }
      }
    } catch (urlError) {
      DEBUG.warn('ScriptApp.getService().getUrl()でエラー', urlError);
    }
    
    // 方法3: フォールバック - スクリプトIDを返す（不完全だが動作する）
    const scriptId = ScriptApp.getScriptId();
    DEBUG.warn('デプロイID取得失敗、スクリプトIDでフォールバック', { scriptId: scriptId.substring(0, 20) + '...' });
    return scriptId;
    
  } catch (error) {
    DEBUG.error('デプロイID取得でエラー', error);
    return ScriptApp.getScriptId(); // 最終フォールバック
  }
}

/**
 * URL詳細分析関数
 * @param {string} url - アクセスされたURL
 * @param {string} userAgent - User-Agent文字列
 * @returns {Object} URL分析結果
 */
function analyzeUrl(url, userAgent) {
  const analysis = {
    originalUrl: url,
    userAgent: userAgent,
    hasUserPath: false,
    userNumber: null,
    scriptId: null,
    cleanUrl: null,
    urlParts: {}
  };
  
  if (!url || url === 'unknown') {
    return analysis;
  }
  
  // URLを分解
  try {
    const urlObj = new URL(url);
    analysis.urlParts = {
      protocol: urlObj.protocol,
      host: urlObj.host,
      pathname: urlObj.pathname,
      search: urlObj.search,
      hash: urlObj.hash
    };
    
    // /u/X/ パターンをチェック
    const userPathMatch = url.match(/\/u\/(\d+)\//);
    if (userPathMatch) {
      analysis.hasUserPath = true;
      analysis.userNumber = userPathMatch[1];
    }
    
    // スクリプトIDを抽出
    const scriptIdMatch = url.match(/\/macros\/(?:u\/\d+\/)?s\/([^\/]+)\//);
    if (scriptIdMatch) {
      analysis.scriptId = scriptIdMatch[1];
    }
    
    // クリーンなURL（/u/X/なし）を生成
    if (analysis.hasUserPath && analysis.scriptId) {
      analysis.cleanUrl = url.replace(/\/u\/\d+\//, '/');
    } else {
      analysis.cleanUrl = url;
    }
    
  } catch (error) {
    analysis.error = error.toString();
  }
  
  return analysis;
}

/**
 * QRスキャナー/WebViewを検出する関数
 * @param {string} userAgent - User-Agent文字列
 * @returns {boolean} QRスキャナーの可能性がある場合true
 */
function detectQrScanner(userAgent) {
  if (!userAgent || userAgent === 'unknown') return false;
  
  const ua = userAgent.toLowerCase();
  
  // QRスキャナーアプリやWebViewのパターン
  const qrScannerPatterns = [
    // iOS
    'ios.*webview',
    'cfnetwork',
    'mobile.*webkit.*version/.*safari',
    
    // Android
    'android.*webview',
    'android.*chrome.*wv',
    'android.*version.*chrome',
    
    // 一般的なQRスキャナーアプリ
    'qr',
    'scanner',
    'camera',
    
    // WebView系
    'webview',
    'embedded',
    'inapp',
    
    // 特定のアプリ
    'line',
    'twitter',
    'facebook',
    'instagram',
    'wechat'
  ];
  
  return qrScannerPatterns.some(pattern => ua.includes(pattern));
}

/**
 * GET リクエスト処理
 */
function doGet(e) {
  // User-Agent情報を取得・解析
  const userAgent = e.request ? e.request.headers['User-Agent'] : 'unknown';
  const referer = e.request ? e.request.headers['Referer'] : 'none';
  const requestUrl = e.request ? e.request.url : 'unknown';
  
  // QRスキャナー/WebViewの検出パターン
  const isQrScanner = detectQrScanner(userAgent);
  const accessInfo = {
    userAgent: userAgent,
    referer: referer,
    requestUrl: requestUrl,
    isQrScanner: isQrScanner,
    timestamp: new Date().toISOString()
  };
  
  DEBUG.log('doGet関数呼び出し', { 
    parameters: e.parameter,
    accessInfo: accessInfo
  });
  
  // 詳細URL分析
  const urlAnalysis = analyzeUrl(requestUrl, userAgent);
  DEBUG.log('📊 URL詳細分析', urlAnalysis);
  
  // QRスキャナーからのアクセスの場合、特別ログ
  if (isQrScanner) {
    DEBUG.log('🔍 QRスキャナーからのアクセス検出', {
      ...accessInfo,
      urlAnalysis: urlAnalysis
    });
  }
  
  const page = e.parameter.page || 'admin';
  const sessionId = e.parameter.sessionId;
  
  // デプロイIDを取得
  const deploymentId = getDeploymentId();
  DEBUG.log('デプロイID取得', { deploymentId: deploymentId.substring(0, 20) + '...' });
  
  DEBUG.log('ページルーティング', { page, sessionId });
  
  try {
    switch (page) {
      case 'admin':
        DEBUG.log('admin.html を呼び出し');
        const adminTemplate = HtmlService.createTemplateFromFile('admin');
        adminTemplate.deploymentId = deploymentId;
        return adminTemplate.evaluate()
          .setTitle('集合知AI セッション管理')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
          
      case 'templates':
        DEBUG.log('templates.html を呼び出し');
        return HtmlService.createTemplateFromFile('templates').evaluate()
          .setTitle('質問テンプレート管理')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
          
        case 'test':
          DEBUG.log('テストページを呼び出し');
          return HtmlService.createHtmlOutput(`
            <h1>テストページ</h1>
            <p>現在時刻: ${new Date().toLocaleString('ja-JP')}</p>
            <p>パラメータ: ${JSON.stringify(e.parameter)}</p>
            <p>GASが正常に動作しています。</p>
            <p><a href="?page=session&sessionId=751f15ed-944f-4aa9-a03a-ab1e733fcf4b">セッションページへ（元版）</a></p>
            <p><a href="?page=session-light&sessionId=751f15ed-944f-4aa9-a03a-ab1e733fcf4b">セッションページへ（軽量版）</a></p>
            <p><a href="?page=session-debug&sessionId=751f15ed-944f-4aa9-a03a-ab1e733fcf4b">デバッグセッションページへ</a></p>
          `);
          
        case 'session-debug':
          DEBUG.log('session-debug.html を呼び出し', { sessionId });
          return HtmlService.createTemplateFromFile('session-debug').evaluate()
            .setTitle('セッション - デバッグ版')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
            
        case 'session-light':
          DEBUG.log('session-light.html を呼び出し', { sessionId });
          if (!sessionId) {
            DEBUG.error('軽量版: セッションIDが未指定');
            return HtmlService.createHtmlOutput(`
              <h1>エラー: セッションIDが指定されていません</h1>
              <p>正しいURL形式: ?page=session-light&sessionId=YOUR_SESSION_ID</p>
              <p>現在のパラメータ: ${JSON.stringify(e.parameter)}</p>
            `);
          }
          
          const lightTemplate = HtmlService.createTemplateFromFile('session-light');
          lightTemplate.sessionId = sessionId;
          lightTemplate.pageParams = JSON.stringify(e.parameter);
          return lightTemplate.evaluate()
            .setTitle('集合知AI セッション - 軽量版')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
          
        case 'session':
          DEBUG.log('session.html を呼び出し', { sessionId });
          if (!sessionId) {
            DEBUG.error('セッションIDが未指定');
            return HtmlService.createHtmlOutput(`
              <h1>エラー: セッションIDが指定されていません</h1>
              <p>正しいURL形式: ?page=session&sessionId=YOUR_SESSION_ID</p>
              <p>現在のパラメータ: ${JSON.stringify(e.parameter)}</p>
            `);
          }
          const sessionTemplate = HtmlService.createTemplateFromFile('session');
          sessionTemplate.sessionId = sessionId;
          sessionTemplate.pageParams = JSON.stringify(e.parameter);
          return sessionTemplate.evaluate()
            .setTitle('集合知AI セッション')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
          
          
      case 'monitor':
        DEBUG.log('monitor.html を呼び出し', { sessionId });
        const monitorSessionId = e.parameter.sessionId;
        if (!monitorSessionId) {
          DEBUG.error('監視画面: セッションIDが未指定');
          return HtmlService.createHtmlOutput('<h1>エラー: セッションIDが指定されていません</h1>');
        }
        const monitorTemplate = HtmlService.createTemplateFromFile('monitor');
        monitorTemplate.sessionId = monitorSessionId;
        monitorTemplate.deploymentId = deploymentId;
        return monitorTemplate.evaluate()
          .setTitle('セッション監視')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
          
      case 'results':
        DEBUG.log('results.html を呼び出し', { sessionId });
        const resultsSessionId = e.parameter.sessionId;
        if (!resultsSessionId) {
          DEBUG.error('結果画面: セッションIDが未指定');
          return HtmlService.createHtmlOutput('<h1>エラー: セッションIDが指定されていません</h1>');
        }
        const resultsTemplate = HtmlService.createTemplateFromFile('results');
        resultsTemplate.sessionId = resultsSessionId;
        resultsTemplate.deploymentId = deploymentId;
        return resultsTemplate.evaluate()
          .setTitle('セッション結果')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
          
      default:
        DEBUG.warn('未知のページ要求', { page });
        return HtmlService.createHtmlOutput('<h1>404: ページが見つかりません</h1>');
    }
  } catch (error) {
    DEBUG.error('doGet処理エラー', error);
    return HtmlService.createHtmlOutput(`<h1>システムエラー: ${error.message}</h1><p>詳細: ${error.toString()}</p>`);
  }
}

/**
 * POST リクエスト処理
 */
function doPost(e) {
  const action = e.parameter.action;
  const data = JSON.parse(e.parameter.data || '{}');
  
  try {
    switch (action) {
      case 'createSession':
        return ContentService.createTextOutput(
          JSON.stringify(createSession(data.theme, data.questionMode, data.templateId, data.customQuestions))
        ).setMimeType(ContentService.MimeType.JSON);
        
      case 'getTemplates':
        return ContentService.createTextOutput(
          JSON.stringify(getQuestionTemplates())
        ).setMimeType(ContentService.MimeType.JSON);
        
      case 'saveTemplate':
        return ContentService.createTextOutput(
          JSON.stringify({
            templateId: saveQuestionTemplate(data.name, data.theme, data.questions),
            success: true
          })
        ).setMimeType(ContentService.MimeType.JSON);
        
      case 'deleteTemplate':
        const deleteResult = deleteQuestionTemplate(data.templateId);
        return ContentService.createTextOutput(
          JSON.stringify(deleteResult)
        ).setMimeType(ContentService.MimeType.JSON);
        
      default:
        throw new Error('Invalid action: ' + action);
    }
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({error: error.toString()})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 質問テンプレートを削除
 * @param {string} templateId - 削除するテンプレートID
 * @returns {Object} 削除結果
 */
function deleteQuestionTemplate(templateId) {
  try {
    if (!templateId) {
      throw new Error('テンプレートIDが指定されていません');
    }
    
    const sheet = getSheet(SHEET_NAMES.QUESTION_TEMPLATES);
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    
    // 削除対象の行を特定（逆順で処理するため）
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === templateId) {
        rowsToDelete.push(i + 1); // シートの行番号は1から始まる
      }
    }
    
    if (rowsToDelete.length === 0) {
      return { success: false, message: 'テンプレートが見つかりません' };
    }
    
    // 行を削除（逆順で削除）
    rowsToDelete.forEach(rowNumber => {
      sheet.deleteRow(rowNumber);
    });
    
    return { 
      success: true, 
      message: 'テンプレートを削除しました',
      deletedRows: rowsToDelete.length
    };
    
  } catch (error) {
    return { 
      success: false, 
      message: 'テンプレート削除エラー: ' + error.toString() 
    };
  }
}

// ============================================================================
// 参加者セッション機能
// ============================================================================

/**
 * セッション情報を取得
 * @param {string} sessionId - セッションID
 * @returns {Object} セッション情報
 */
function getSessionInfo(sessionId) {
  try {
    // デバッグ: 渡ってきているsessionIdを出力
    DEBUG.log('🔍 getSessionInfo デバッグ開始', { 
      sessionId: sessionId,
      sessionIdType: typeof sessionId,
      sessionIdLength: sessionId ? sessionId.length : 'undefined'
    });
    
    if (!sessionId) {
      return { success: false, message: 'セッションIDが指定されていません' };
    }
    
    const sheet = getSheet(SHEET_NAMES.QUESTIONS);
    const spreadsheetId = sheet.getParent().getId();
    const spreadsheetName = sheet.getParent().getName();
    const sheetName = sheet.getName();
    
    // デバッグ: スプレッドシート情報を出力
    DEBUG.log('📊 スプレッドシート情報', {
      spreadsheetId: spreadsheetId,
      spreadsheetName: spreadsheetName,
      sheetName: sheetName,
      sheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheet.getSheetId()}`
    });
    
    const data = sheet.getDataRange().getValues();
    DEBUG.log('📋 シートデータ情報', {
      totalRows: data.length,
      dataHeader: data[0] || [],
      sampleData: data.slice(0, 3) // ヘッダー含む最初の3行
    });
    
    let theme = '';
    const questions = ['', '', ''];
    const questionTypes = ['fixed_1', 'fixed_2', 'free_discussion'];
    
    // セッションIDに該当する質問を検索
    let foundRows = 0;
    DEBUG.log('🔎 セッションID検索開始', { 
      targetSessionId: sessionId,
      searchingInRows: data.length - 1 
    });
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowSessionId = row[0];
      
      // 各行の詳細をデバッグ出力（最初の5行のみ）
      if (i <= 5) {
        DEBUG.log(`📄 行${i}データ`, {
          sessionId: rowSessionId,
          theme: row[1],
          questionType: row[2],
          questionText: row[3] ? row[3].substring(0, 50) + '...' : 'empty'
        });
      }
      
      if (rowSessionId === sessionId) {
        foundRows++;
        theme = row[1]; // theme列
        const questionType = row[2]; // question_type列
        const questionText = row[3]; // question_text列
        
        DEBUG.log('✅ マッチした行を発見', {
          rowIndex: i,
          sessionId: rowSessionId,
          theme: theme,
          questionType: questionType,
          questionText: questionText ? questionText.substring(0, 100) + '...' : 'empty',
          questionTextLength: questionText ? questionText.length : 0
        });
        
        const typeIndex = questionTypes.indexOf(questionType);
        if (typeIndex !== -1) {
          questions[typeIndex] = questionText;
          DEBUG.log('📝 質問を配列に設定', {
            questionType: questionType,
            typeIndex: typeIndex,
            questionSet: questionText ? 'success' : 'failed (empty text)'
          });
        } else {
          DEBUG.warn('⚠️ 不明な質問タイプ', { questionType: questionType });
        }
      }
    }
    
    DEBUG.log('🔍 検索結果サマリー', {
      foundRows: foundRows,
      theme: theme,
      questionsFound: questions.filter(q => q && q.trim() !== '').length,
      allQuestions: questions.map((q, i) => ({
        type: questionTypes[i],
        hasContent: !!q && q.trim() !== '',
        length: q ? q.length : 0,
        preview: q ? q.substring(0, 50) + '...' : 'empty'
      }))
    });
    
    if (!theme) {
      DEBUG.error('❌ セッションが見つかりません', { 
        sessionId: sessionId,
        foundRows: foundRows,
        totalDataRows: data.length - 1
      });
      return { success: false, message: 'セッションが見つかりません' };
    }
    
    // 全ての質問が取得できているかチェック
    const emptyQuestions = questions.filter(q => !q || q.trim() === '');
    if (emptyQuestions.length > 0) {
      DEBUG.error('❌ 質問が不完全', { 
        sessionId, 
        theme, 
        questions: questions.map((q, i) => ({
          type: questionTypes[i],
          content: q || 'EMPTY',
          isEmpty: !q || q.trim() === ''
        })),
        emptyCount: emptyQuestions.length,
        emptyIndexes: questions.map((q, i) => !q || q.trim() === '' ? i : null).filter(i => i !== null)
      });
      return { 
        success: false, 
        message: `セッションの質問データが不完全です（${emptyQuestions.length}問が不足）` 
      };
    }
    
    // 成功時の最終結果をデバッグ出力
    DEBUG.log('✅ セッション情報取得成功', {
      sessionId: sessionId,
      theme: theme,
      questions: questions.map((q, i) => ({
        type: questionTypes[i],
        length: q.length,
        preview: q.substring(0, 100) + (q.length > 100 ? '...' : ''),
        fullContent: q // 質問の完全な内容
      })),
      totalQuestions: questions.length,
      completedAt: new Date().toISOString()
    });
    
    return {
      success: true,
      sessionId: sessionId,
      theme: theme,
      questions: questions
    };
    
  } catch (error) {
    return { 
      success: false, 
      message: 'セッション情報取得エラー: ' + error.toString() 
    };
  }
}

/**
 * 参加者の回答に対するAIフィードバックを生成
 * @param {Object} data - 回答データ
 * @returns {Object} フィードバック結果
 */
function getAiFeedbackForAnswer(data) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const endpoint = PropertiesService.getScriptProperties().getProperty('GEMINI_API_ENDPOINT') || 
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent';
  
  if (!apiKey) {
    return { success: false, message: 'Gemini API Keyが設定されていません' };
  }
  
  const prompt = `
セッションテーマ: ${data.theme}
質問: ${data.question}
参加者の回答: ${data.answer}

この回答に対して、以下の観点から建設的なフィードバックを150文字以内で提供してください：
1. 回答の良い点を1つ指摘
2. より深く考えるための追加の視点を1つ提案
3. 他の参加者との議論につながる要素を1つ提示

フィードバックは参加者が前向きに感じられるような温かい口調で記述してください。
`;

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 800,
    }
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-goog-api-key': apiKey
    },
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    const responseData = JSON.parse(response.getContentText());
    const feedback = responseData.candidates[0].content.parts[0].text;
    
    return {
      success: true,
      feedback: feedback.trim()
    };
    
  } catch (error) {
    // フォールバック応答
    return {
      success: true,
      feedback: `ご回答ありがとうございます！「${data.answer.substring(0, 30)}...」という視点は興味深いですね。他の参加者の意見と合わせて、より深い議論につながりそうです。`
    };
  }
}

/**
 * セッション結果を保存
 * @param {Object} data - セッション完了データ
 * @returns {Object} 保存結果
 */
function saveSessionResults(data) {
  try {
    DEBUG.log('🔍 [SAVE] saveSessionResults呼び出し', {
      sessionId: data.sessionId,
      participantId: data.participantId,
      participantName: data.participantName,
      answersCount: data.answers ? data.answers.length : 0,
      aiResponsesCount: data.aiResponses ? data.aiResponses.length : 0,
      answersDetail: data.answers ? data.answers.map((answer, i) => ({
        index: i,
        hasAnswer: !!answer && answer.trim() !== '',
        answerLength: answer ? answer.length : 0,
        preview: answer ? answer.substring(0, 30) + '...' : 'empty'
      })) : []
    });
    
    const sheet = getSheet(SHEET_NAMES.RESULTS);
    const questionTypes = ['fixed_1', 'fixed_2', 'free_discussion'];
    
    // サーバーサイドで現在時刻を生成（文字列化してDate問題回避）
    const timestamp = new Date().toISOString();
    
    DEBUG.log('📊 [SAVE] セッション結果保存開始', { 
      timestamp: timestamp,
      questionTypes: questionTypes
    });
    
    let savedCount = 0;
    let skippedCount = 0;
    
    // 各質問の回答を1行ずつ保存
    data.answers.forEach((answer, index) => {
      const questionType = questionTypes[index];
      const aiResponse = data.aiResponses && data.aiResponses[index] ? data.aiResponses[index] : '';
      
      DEBUG.log(`📝 [SAVE] 質問${index + 1}(${questionType})処理`, {
        hasAnswer: !!answer && answer.trim() !== '',
        answerLength: answer ? answer.length : 0,
        hasAiResponse: !!aiResponse && aiResponse.trim() !== '',
        aiResponseLength: aiResponse ? aiResponse.length : 0
      });
      
      if (answer && answer.trim()) {
        try {
          sheet.appendRow([
            data.sessionId,
            data.participantId,
            `${data.sessionId}_${questionType}`, // question_id
            questionType,
            answer.trim(),
            aiResponse,
            timestamp
          ]);
          savedCount++;
          DEBUG.log(`✅ [SAVE] 質問${index + 1}保存成功`, {
            questionType: questionType,
            answerPreview: answer.substring(0, 50) + '...',
            aiResponsePreview: aiResponse ? aiResponse.substring(0, 50) + '...' : 'empty'
          });
        } catch (rowError) {
          DEBUG.error(`❌ [SAVE] 質問${index + 1}保存エラー`, {
            questionType: questionType,
            error: rowError.toString()
          });
        }
      } else {
        skippedCount++;
        DEBUG.warn(`⚠️ [SAVE] 質問${index + 1}スキップ`, {
          questionType: questionType,
          reason: answer ? 'empty after trim' : 'null or undefined'
        });
      }
    });
    
    DEBUG.log('📊 [SAVE] セッション結果保存完了', {
      savedCount: savedCount,
      skippedCount: skippedCount,
      totalAnswers: data.answers ? data.answers.length : 0
    });
    
    return {
      success: true,
      message: `セッション結果を保存しました（保存: ${savedCount}件、スキップ: ${skippedCount}件）`,
      participantId: data.participantId,
      savedCount: savedCount,
      skippedCount: skippedCount
    };
    
  } catch (error) {
    DEBUG.error('❌ [SAVE] セッション結果保存エラー', {
      error: error.toString(),
      stack: error.stack,
      sessionId: data ? data.sessionId : 'unknown'
    });
    return {
      success: false,
      message: 'セッション結果保存エラー: ' + error.toString()
    };
  }
}

/**
 * セッション結果を取得
 * @param {string} sessionId - セッションID
 * @returns {Object} セッション結果
 */
function getSessionResults(sessionId) {
  try {
    if (!sessionId) {
      return { success: false, message: 'セッションIDが指定されていません' };
    }
    
    // セッション情報を取得
    const sessionInfo = getSessionInfo(sessionId);
    if (!sessionInfo.success) {
      return { success: false, message: 'セッション情報の取得に失敗しました' };
    }
    
    // 結果データを取得
    const resultsSheet = getSheet(SHEET_NAMES.RESULTS);
    const resultsData = resultsSheet.getDataRange().getValues();
    
    const responses = [];
    const participantIds = new Set();
    
    // セッションIDに該当する結果を検索
    for (let i = 1; i < resultsData.length; i++) {
      const row = resultsData[i];
      if (row[0] === sessionId) {
        responses.push({
          session_id: row[0],
          participant_id: row[1],
          question_id: row[2],
          question_type: row[3],
          user_input: row[4],
          ai_response: row[5],
          timestamp: row[6]
        });
        participantIds.add(row[1]);
      }
    }
    
    // 質問別に回答を整理
    const responsesByQuestion = {};
    const questionTypes = ['fixed_1', 'fixed_2', 'free_discussion'];
    
    questionTypes.forEach(type => {
      responsesByQuestion[type] = responses.filter(r => r.question_type === type);
    });
    
    // サマリー情報を取得（存在する場合）
    let summary = null;
    try {
      const summarySheet = getSheet(SHEET_NAMES.SUMMARY);
      const summaryData = summarySheet.getDataRange().getValues();
      
      for (let i = 1; i < summaryData.length; i++) {
        const row = summaryData[i];
        if (row[0] === sessionId) {
          summary = {
            consensusPoints: row[3] ? row[3].split('\n').filter(p => p.trim()) : [],
            divergentPoints: row[4] ? row[4].split('\n').filter(p => p.trim()) : [],
            keyInsights: row[5] || '',
            fullSummary: row[6] || ''
          };
          break;
        }
      }
    } catch (error) {
      console.log('サマリー取得エラー:', error);
    }
    
    return {
      success: true,
      sessionId: sessionId,
      theme: sessionInfo.theme,
      questions: sessionInfo.questions,
      participantCount: participantIds.size,
      responseCount: responses.length,
      responsesByQuestion: responsesByQuestion,
      summary: summary
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'セッション結果取得エラー: ' + error.toString()
    };
  }
}

/**
 * AI分析によるセッションサマリーを生成
 * @param {string} sessionId - セッションID
 * @returns {Object} 分析結果
 */
function generateSessionSummary(sessionId) {
  try {
    const sessionResults = getSessionResults(sessionId);
    if (!sessionResults.success) {
      return { success: false, message: 'セッションデータの取得に失敗しました' };
    }
    
    if (sessionResults.responseCount === 0) {
      return { success: false, message: '分析するデータがありません' };
    }
    
    // 全回答を結合
    const allResponses = [];
    Object.keys(sessionResults.responsesByQuestion).forEach(questionType => {
      const questionNumber = getQuestionNumberFromType(questionType);
      const questionText = sessionResults.questions[questionNumber - 1];
      
      sessionResults.responsesByQuestion[questionType].forEach(response => {
        allResponses.push({
          question: questionText,
          answer: response.user_input,
          questionType: questionType
        });
      });
    });
    
    // Gemini APIで分析
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    const endpoint = PropertiesService.getScriptProperties().getProperty('GEMINI_API_ENDPOINT') || 
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent';
    
    if (!apiKey) {
      return { success: false, message: 'Gemini API Keyが設定されていません' };
    }
    
    const prompt = buildAnalysisPrompt(sessionResults.theme, allResponses);
    
    const payload = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }],
      generationConfig: {
        temperature: 0.3,
        maxOutputTokens: 2000,
      }
    };
    
    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-goog-api-key': apiKey
      },
      payload: JSON.stringify(payload)
    };
    
    const response = UrlFetchApp.fetch(endpoint, options);
    const data = JSON.parse(response.getContentText());
    const analysisText = data.candidates[0].content.parts[0].text;
    
    // 分析結果をパース
    const analysis = parseAnalysisResult(analysisText);
    
    // Summaryシートに保存
    saveSummaryToSheet(sessionId, sessionResults.theme, sessionResults.participantCount, analysis);
    
    return {
      success: true,
      consensusPoints: analysis.consensus,
      divergentPoints: analysis.divergent,
      keyInsights: analysis.insights,
      fullSummary: analysisText
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'AI分析エラー: ' + error.toString()
    };
  }
}

/**
 * 分析用プロンプトを構築
 */
function buildAnalysisPrompt(theme, responses) {
  let prompt = `
セッションテーマ: ${theme}
参加者数: ${new Set(responses.map(r => r.participant_id)).size}人
回答数: ${responses.length}件

以下の回答を分析し、合意点と多様な意見を抽出してください：

`;

  responses.forEach((response, index) => {
    prompt += `
【回答${index + 1}】
質問: ${response.question}
回答: ${response.answer}
`;
  });

  prompt += `

以下の形式で分析結果を出力してください：

【合意点・共通認識】
- 参加者間で共通している考えや意見
- 多くの人が同意している要素
- 基本的な前提や価値観の共有部分

【多様な意見・分散点】  
- 参加者間で意見が分かれている部分
- 異なる視点や観点
- 個性的で興味深い発想

【重要な洞察】
- セッション全体から得られる気づき
- 今後の議論につながる要素
- 意思決定に役立つポイント

各項目は簡潔な箇条書きで、1項目につき50文字以内で記述してください。
`;

  return prompt;
}

/**
 * 分析結果をパース
 */
function parseAnalysisResult(analysisText) {
  const sections = {
    consensus: [],
    divergent: [],
    insights: ''
  };
  
  try {
    // 正規表現で各セクションを抽出
    const consensusMatch = analysisText.match(/【合意点・共通認識】([\s\S]*?)【多様な意見・分散点】/);
    const divergentMatch = analysisText.match(/【多様な意見・分散点】([\s\S]*?)【重要な洞察】/);
    const insightsMatch = analysisText.match(/【重要な洞察】([\s\S]*?)$/);
    
    if (consensusMatch) {
      sections.consensus = consensusMatch[1]
        .split('\n')
        .filter(line => line.trim().startsWith('-'))
        .map(line => line.trim().substring(1).trim());
    }
    
    if (divergentMatch) {
      sections.divergent = divergentMatch[1]
        .split('\n')
        .filter(line => line.trim().startsWith('-'))
        .map(line => line.trim().substring(1).trim());
    }
    
    if (insightsMatch) {
      sections.insights = insightsMatch[1].trim();
    }
    
  } catch (error) {
    console.log('分析結果パースエラー:', error);
    // フォールバック: シンプルな抽出
    sections.consensus = ['参加者の意見を分析中です'];
    sections.divergent = ['多様な視点が確認されました'];
    sections.insights = '詳細な分析結果を準備中です';
  }
  
  return sections;
}

/**
 * サマリーをSummaryシートに保存
 */
function saveSummaryToSheet(sessionId, theme, participantCount, analysis) {
  try {
    const sheet = getSheet(SHEET_NAMES.SUMMARY);
    const consensusText = analysis.consensus.join('\n');
    const divergentText = analysis.divergent.join('\n');
    
    // 既存のサマリーがあれば更新、なければ新規作成
    const data = sheet.getDataRange().getValues();
    let rowFound = false;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === sessionId) {
        // 更新
        sheet.getRange(i + 1, 3, 1, 5).setValues([[
          participantCount,
          consensusText,
          divergentText,
          analysis.insights,
          new Date()
        ]]);
        rowFound = true;
        break;
      }
    }
    
    if (!rowFound) {
      // 新規作成
      sheet.appendRow([
        sessionId,
        theme,
        participantCount,
        consensusText,
        divergentText,
        analysis.insights,
        new Date()
      ]);
    }
    
  } catch (error) {
    console.log('サマリー保存エラー:', error);
  }
}

/**
 * 質問タイプから番号を取得
 */
function getQuestionNumberFromType(questionType) {
  const mapping = {
    'fixed_1': 1,
    'fixed_2': 2,
    'free_discussion': 3
  };
  return mapping[questionType] || 1;
}

// ============================================================================
// デバッグ支援関数
// ============================================================================

/**
 * デバッグモードの切り替え
 * @param {boolean} enabled - デバッグモードを有効にするか
 */
function setDebugMode(enabled = true) {
  DEBUG.log(`デバッグモードを${enabled ? '有効' : '無効'}に設定`);
  PropertiesService.getScriptProperties().setProperty('DEBUG_MODE', enabled.toString());
  return { success: true, debugMode: enabled };
}

/**
 * デバッグログの表示
 * @param {number} limit - 表示する行数（デフォルト20）
 */
function showDebugLogs(limit = 20) {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      console.log('スプレッドシートIDが設定されていません');
      return;
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const logSheet = spreadsheet.getSheetByName('DebugLog');
    
    if (!logSheet) {
      console.log('DebugLogシートが存在しません');
      return;
    }
    
    const data = logSheet.getDataRange().getValues();
    const logs = data.slice(-limit);
    
    console.log(`=== 最新${logs.length}件のデバッグログ ===`);
    logs.forEach((row, index) => {
      const [timestamp, level, message, data, caller] = row;
      console.log(`[${index + 1}] ${timestamp} [${level}] ${message}`);
      if (data) console.log(`    データ: ${data}`);
      if (caller) console.log(`    呼び出し元: ${caller}`);
      console.log('');
    });
    
    return { success: true, logCount: logs.length };
  } catch (error) {
    console.error('デバッグログ表示エラー:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * デバッグログのクリア
 */
function clearDebugLogs() {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      return { success: false, message: 'スプレッドシートIDが設定されていません' };
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const logSheet = spreadsheet.getSheetByName('DebugLog');
    
    if (!logSheet) {
      return { success: false, message: 'DebugLogシートが存在しません' };
    }
    
    // ヘッダー行以外をクリア
    const lastRow = logSheet.getLastRow();
    if (lastRow > 1) {
      logSheet.getRange(2, 1, lastRow - 1, logSheet.getLastColumn()).clear();
    }
    
    DEBUG.log('デバッグログをクリアしました');
    return { success: true, message: 'デバッグログをクリアしました' };
  } catch (error) {
    console.error('デバッグログクリアエラー:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * システム状態の詳細表示（デバッグ用）
 */
function debugSystemStatus() {
  DEBUG.log('システム状態詳細確認開始');
  console.log('=== システム状態詳細 ===');
  
  // スクリプトプロパティ確認
  const properties = PropertiesService.getScriptProperties().getProperties();
  console.log('スクリプトプロパティ:', Object.keys(properties));
  
  // API Key確認
  const apiKey = properties.GEMINI_API_KEY;
  console.log('API Key設定:', apiKey ? `設定済み (${apiKey.length}文字)` : '未設定');
  
  // スプレッドシート確認
  const spreadsheetId = properties.SPREADSHEET_ID;
  console.log('スプレッドシートID:', spreadsheetId ? `設定済み (${spreadsheetId.substring(0, 10)}...)` : '未設定');
  
  if (spreadsheetId) {
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      console.log('スプレッドシート名:', spreadsheet.getName());
      
      const sheets = spreadsheet.getSheets();
      console.log('既存シート:', sheets.map(s => s.getName()));
      
      // 各シートのデータ量確認
      Object.values(SHEET_NAMES).forEach(sheetName => {
        try {
          const sheet = spreadsheet.getSheetByName(sheetName);
          if (sheet) {
            console.log(`${sheetName}: ${sheet.getLastRow() - 1}行のデータ`);
          } else {
            console.log(`${sheetName}: シート未作成`);
          }
        } catch (e) {
          console.log(`${sheetName}: アクセスエラー`);
        }
      });
    } catch (error) {
      console.error('スプレッドシートアクセスエラー:', error);
    }
  }
  
  DEBUG.log('システム状態詳細確認完了');
  return { success: true };
}

/**
 * セッション監視用データを取得
 * @param {string} sessionId - セッションID
 * @returns {Object} 監視データ
 */
function getSessionMonitorData(sessionId) {
  try {
    DEBUG.log('🔍 [MONITOR] getSessionMonitorData開始', { 
      sessionId: sessionId,
      sessionIdType: typeof sessionId,
      sessionIdLength: sessionId ? sessionId.length : 'undefined'
    });
    
    if (!sessionId) {
      return { success: false, message: 'セッションIDが指定されていません' };
    }
    
    // 直接Resultsシートからデータを取得
    const sheet = getSheet(SHEET_NAMES.RESULTS);
    const spreadsheetId = sheet.getParent().getId();
    const spreadsheetName = sheet.getParent().getName();
    const sheetName = sheet.getName();
    
    DEBUG.log('📊 [MONITOR] Resultsシート情報', {
      spreadsheetId: spreadsheetId,
      spreadsheetName: spreadsheetName,
      sheetName: sheetName,
      sheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheet.getSheetId()}`
    });
    
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0]; // ['session_id', 'participant_id', 'question_id', 'question_type', 'user_input', 'ai_response', 'timestamp']
    
    DEBUG.log('📋 [MONITOR] Resultsシートデータ情報', {
      totalRows: allData.length,
      headers: headers,
      sampleData: allData.slice(0, 3) // ヘッダー含む最初の3行
    });
    
    // セッションIDに該当するデータをフィルタリング
    const sessionResponses = [];
    let matchedRows = 0;
    
    DEBUG.log('🔎 [MONITOR] セッションID検索開始', { 
      targetSessionId: sessionId,
      searchingInRows: allData.length - 1 
    });
    
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const rowSessionId = row[0];
      
      // 最初の5行の詳細をデバッグ出力
      if (i <= 5) {
        DEBUG.log(`📄 [MONITOR] 行${i}データ`, {
          sessionId: rowSessionId,
          participantId: row[1],
          questionType: row[3]
          // timestampは削除（Date問題回避）
        });
      }
      
      if (rowSessionId === sessionId) { // session_idが一致
        matchedRows++;
        const responseData = {
          session_id: row[0],
          participant_id: row[1],
          question_id: row[2],
          question_type: row[3],
          user_input: row[4],
          ai_response: row[5]
          // timestampは削除（Date問題回避）
        };
        sessionResponses.push(responseData);
        
        if (matchedRows <= 3) { // 最初の3件のマッチを詳細ログ
          DEBUG.log(`✅ [MONITOR] マッチした行${i}`, responseData);
        }
      }
    }
    
    DEBUG.log('🔍 [MONITOR] 検索結果サマリー', { 
      matchedRows: matchedRows,
      sessionResponsesCount: sessionResponses.length,
      targetSessionId: sessionId
    });
    
    // セッション基本情報を取得
    DEBUG.log('📋 [MONITOR] セッション基本情報取得開始', { sessionId });
    const sessionInfo = getSessionInfo(sessionId);
    const theme = sessionInfo.success ? sessionInfo.theme : 'テーマ取得失敗';
    
    DEBUG.log('📋 [MONITOR] セッション基本情報取得結果', { 
      sessionInfoSuccess: sessionInfo.success,
      theme: theme,
      sessionInfoMessage: sessionInfo.message || 'N/A'
    });
    
    // 質問別回答数をカウント
    const questionCounts = {
      fixed_1: 0,
      fixed_2: 0,
      free_discussion: 0
    };
    
    sessionResponses.forEach(response => {
      if (questionCounts.hasOwnProperty(response.question_type)) {
        questionCounts[response.question_type]++;
      }
    });
    
    DEBUG.log('📊 [MONITOR] 質問別回答数', questionCounts);
    
    // 時刻なしのシンプルな回答リスト（最新20件）
    const recentActivities = sessionResponses.slice(-20).reverse(); // 最後の20件を逆順で
    
    DEBUG.log('📄 [MONITOR] 最新アクティビティ', { 
      totalResponses: sessionResponses.length,
      recentActivitiesCount: recentActivities.length
    });
    
    // 参加者別の進捗情報を作成（時刻なし）
    const participantsMap = new Map();
    sessionResponses.forEach(response => {
      const participantId = response.participant_id;
      if (!participantId) {
        DEBUG.warn('⚠️ [MONITOR] participant_idが空', { response });
        return;
      }
      
      if (!participantsMap.has(participantId)) {
        participantsMap.set(participantId, {
          id: participantId,
          responseCount: 0
          // lastActivityは削除（Date問題回避）
        });
      }
      
      const participant = participantsMap.get(participantId);
      participant.responseCount++;
    });
    
    // 参加者リストを配列に変換（回答数でソート）
    const participants = Array.from(participantsMap.values())
      .sort((a, b) => b.responseCount - a.responseCount);
    
    // 統計計算
    const participantCount = participants.length;
    const responseCount = sessionResponses.length;
    
    DEBUG.log('📊 [MONITOR] 最終統計情報', { 
      participantCount, 
      responseCount, 
      questionCounts,
      participantsMapSize: participantsMap.size
    });
    
    // シンプルで確実なデータ構造（Dateオブジェクト完全排除）
    const result = {
      success: true,
      sessionId: sessionId,
      theme: theme,
      questions: sessionInfo.success ? sessionInfo.questions : [],
      participantCount: participantCount,
      responseCount: responseCount,
      questionCounts: questionCounts,
      recentActivities: recentActivities,
      participants: participants
      // lastUpdatedは削除（Date問題回避）
    };
    
    DEBUG.log('✅ [MONITOR] 監視データ作成完了', { 
      resultSuccess: result.success,
      resultKeys: Object.keys(result),
      theme: result.theme,
      participantCount: result.participantCount,
      responseCount: result.responseCount
    });
    
    // 返却直前の最終チェック
    DEBUG.log('🚀 [MONITOR] データ返却直前', {
      resultType: typeof result,
      resultIsNull: result === null,
      resultIsUndefined: result === undefined,
      resultStringLength: JSON.stringify(result).length,
      resultAsStringPreview: JSON.stringify(result).substring(0, 200) + '...',
      willReturnSuccess: result && result.success
    });
    
    return result;
    
  } catch (error) {
    DEBUG.error('❌ [MONITOR] getSessionMonitorData エラー', { 
      error: error.toString(),
      stack: error.stack,
      sessionId: sessionId
    });
    return {
      success: false,
      message: '監視データ取得エラー: ' + error.toString()
    };
  }
}

/**
 * セッション結果を取得（results.html用）
 * @param {string} sessionId - セッションID
 * @returns {Object} セッション結果データ
 */
function getSessionResults(sessionId) {
  try {
    DEBUG.log('🔍 [RESULTS] getSessionResults開始', { 
      sessionId: sessionId,
      sessionIdType: typeof sessionId,
      sessionIdLength: sessionId ? sessionId.length : 'undefined'
    });
    
    if (!sessionId) {
      return { success: false, message: 'セッションIDが指定されていません' };
    }
    
    // セッション基本情報を取得
    const sessionInfo = getSessionInfo(sessionId);
    if (!sessionInfo.success) {
      return { success: false, message: 'セッション情報の取得に失敗しました: ' + sessionInfo.message };
    }
    
    DEBUG.log('📋 [RESULTS] セッション基本情報取得', { 
      theme: sessionInfo.theme,
      questionsCount: sessionInfo.questions ? sessionInfo.questions.length : 0
    });
    
    // Resultsシートからデータを取得
    const sheet = getSheet(SHEET_NAMES.RESULTS);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0]; // ['session_id', 'participant_id', 'question_id', 'question_type', 'user_input', 'ai_response', 'timestamp']
    
    DEBUG.log('📊 [RESULTS] Resultsシート情報', {
      spreadsheetId: sheet.getParent().getId(),
      sheetName: sheet.getName(),
      totalRows: allData.length,
      headers: headers
    });
    
    // セッションに該当する回答データを取得
    const sessionResponses = [];
    let matchedRows = 0;
    
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const rowSessionId = row[0];
      
      if (rowSessionId === sessionId) {
        matchedRows++;
        const responseData = {
          session_id: row[0],
          participant_id: row[1],
          question_id: row[2],
          question_type: row[3],
          user_input: row[4],
          ai_response: row[5]
          // timestampは除外（Date問題回避）
        };
        sessionResponses.push(responseData);
      }
    }
    
    DEBUG.log('🔎 [RESULTS] 回答データ取得結果', { 
      matchedRows: matchedRows,
      sessionResponsesCount: sessionResponses.length
    });
    
    // 質問タイプ別に回答をグループ化（results.html形式に合わせる）
    const questionTypes = ['fixed_1', 'fixed_2', 'free_discussion'];
    const responsesByQuestion = {};
    
    questionTypes.forEach((type, index) => {
      const questionText = sessionInfo.questions && sessionInfo.questions[index] ? sessionInfo.questions[index] : `質問${index + 1}`;
      const answers = sessionResponses.filter(r => r.question_type === type);
      
      responsesByQuestion[type] = answers.map(answer => ({
        participant_id: answer.participant_id,
        user_input: answer.user_input || '',
        ai_response: answer.ai_response || ''
        // timestampは除外（Date問題回避）
      }));
    });
    
    DEBUG.log('📝 [RESULTS] 質問別グループ化完了', { 
      fixed_1_count: responsesByQuestion.fixed_1.length,
      fixed_2_count: responsesByQuestion.fixed_2.length,
      free_discussion_count: responsesByQuestion.free_discussion.length
    });
    
    // 参加者統計
    const participantIds = [...new Set(sessionResponses.map(r => r.participant_id))];
    const participantStats = participantIds.map(id => {
      const participantAnswers = sessionResponses.filter(r => r.participant_id === id);
      return {
        participantId: id,
        participantNumber: id.replace('participant_', 'P'),
        answerCount: participantAnswers.length,
        completionRate: Math.round((participantAnswers.length / 3) * 100)
      };
    });
    
    DEBUG.log('👥 [RESULTS] 参加者統計完了', { 
      totalParticipants: participantStats.length,
      avgCompletionRate: participantStats.length > 0 ? 
        Math.round(participantStats.reduce((sum, p) => sum + p.completionRate, 0) / participantStats.length) : 0
    });
    
    // 最終結果データ（results.html形式に合わせる）
    const result = {
      success: true,
      sessionId: sessionId,
      theme: sessionInfo.theme,
      questions: sessionInfo.questions,
      responsesByQuestion: responsesByQuestion,
      participantCount: participantStats.length,
      responseCount: sessionResponses.length,
      participantStats: participantStats,
      avgCompletionRate: participantStats.length > 0 ? 
        Math.round(participantStats.reduce((sum, p) => sum + p.completionRate, 0) / participantStats.length) : 0
    };
    
    DEBUG.log('✅ [RESULTS] セッション結果データ作成完了', { 
      resultSuccess: result.success,
      theme: result.theme,
      totalParticipants: result.totalParticipants,
      totalAnswers: result.totalAnswers,
      avgCompletionRate: result.avgCompletionRate
    });
    
    return result;
    
  } catch (error) {
    DEBUG.error('❌ [RESULTS] getSessionResults エラー', { 
      error: error.toString(),
      stack: error.stack,
      sessionId: sessionId
    });
    return { success: false, message: error.toString() };
  }
}