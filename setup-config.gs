/**
 * 初期設定・環境設定関連の関数
 * Google Apps Script のプロパティサービスを使用して設定値を管理
 * セキュリティのため、APIキーやIDは手動でGAS GUIから設定
 */

// 必要なスクリプトプロパティ定数の定義
const REQUIRED_PROPERTIES = {
  // 機密情報（GAS GUIから手動設定）
  GEMINI_API_KEY: {
    description: 'Gemini API Key（Google AI Studioから取得）',
    required: true,
    sensitive: true
  },
  SPREADSHEET_ID: {
    description: 'Google SpreadsheetsのID（URLから取得）',
    required: true,
    sensitive: false
  },
  
  // システム設定（デフォルト値あり）
  SESSION_TIMEOUT_MINUTES: {
    description: 'セッションタイムアウト（分）',
    required: false,
    defaultValue: '5'
  },
  MAX_PARTICIPANTS_PER_SESSION: {
    description: '1セッションあたりの最大参加者数',
    required: false,
    defaultValue: '30'
  },
  GEMINI_API_ENDPOINT: {
    description: 'Gemini APIエンドポイント',
    required: false,
    defaultValue: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent'
  },
  API_REQUEST_TIMEOUT: {
    description: 'APIリクエストタイムアウト（ミリ秒）',
    required: false,
    defaultValue: '30000'
  },
  MAX_API_RETRIES: {
    description: 'API呼び出し最大リトライ回数',
    required: false,
    defaultValue: '3'
  },
  APP_VERSION: {
    description: 'アプリケーションバージョン',
    required: false,
    defaultValue: '1.0.0'
  },
  DEBUG_MODE: {
    description: 'デバッグモードフラグ',
    required: false,
    defaultValue: 'false'
  },
  SYSTEM_TIMEZONE: {
    description: 'システムタイムゾーン',
    required: false,
    defaultValue: 'Asia/Tokyo'
  }
};

/**
 * デフォルト値のみを設定する安全な初期化
 * 機密情報はGAS GUIから手動設定する必要があります
 */
function initializeScriptProperties() {
  const properties = PropertiesService.getScriptProperties();
  
  try {
    const defaultProperties = {};
    
    // デフォルト値があるプロパティのみ設定
    Object.entries(REQUIRED_PROPERTIES).forEach(([key, config]) => {
      if (config.defaultValue && !config.sensitive) {
        defaultProperties[key] = config.defaultValue;
      }
    });
    
    properties.setProperties(defaultProperties);
    
    console.log('✅ デフォルト設定値が設定されました');
    console.log('⚠️  重要: 以下の値をGAS GUIから手動で設定してください:');
    console.log('   - GEMINI_API_KEY: Gemini API Key');
    console.log('   - SPREADSHEET_ID: Google SpreadsheetsのID');
    
    return {
      success: true,
      message: 'デフォルト設定完了。機密情報はGAS GUIから手動設定してください。',
      requiredManualSettings: ['GEMINI_API_KEY', 'SPREADSHEET_ID']
    };
    
  } catch (error) {
    console.error('❌ 初期化エラー:', error);
    return {
      success: false,
      message: '初期化中にエラーが発生しました: ' + error.toString()
    };
  }
}

/**
 * 必要なプロパティの設定状況を表示
 * GAS GUIでの手動設定を支援
 */
function showRequiredProperties() {
  console.log('=== 必要なスクリプトプロパティ ===');
  console.log('GAS GUI（プロジェクト設定 > スクリプトプロパティ）で以下を設定してください:\n');
  
  Object.entries(REQUIRED_PROPERTIES).forEach(([key, config]) => {
    const status = config.required ? '[必須]' : '[オプション]';
    const sensitive = config.sensitive ? '[機密]' : '';
    const defaultVal = config.defaultValue ? `(デフォルト: ${config.defaultValue})` : '';
    
    console.log(`${status}${sensitive} ${key}`);
    console.log(`  説明: ${config.description} ${defaultVal}`);
    console.log('');
  });
  
  console.log('設定手順:');
  console.log('1. GASエディタで「プロジェクト設定」をクリック');
  console.log('2. 「スクリプトプロパティ」セクションを開く');
  console.log('3. 「スクリプトプロパティを追加」で上記のキーと値を設定');
}

/**
 * 現在の設定値を確認・表示
 */
function showCurrentConfiguration() {
  const properties = PropertiesService.getScriptProperties().getProperties();
  
  console.log('=== 現在の設定値 ===');
  
  // API Key（セキュリティのため一部のみ表示）
  const apiKey = properties.GEMINI_API_KEY || '未設定';
  if (apiKey !== '未設定' && apiKey.length > 10) {
    console.log('Gemini API Key:', apiKey.substring(0, 10) + '...' + apiKey.substring(apiKey.length - 4));
  } else {
    console.log('Gemini API Key:', apiKey);
  }
  
  // スプレッドシートID
  const spreadsheetId = properties.SPREADSHEET_ID || '未設定';
  console.log('Spreadsheet ID:', spreadsheetId);
  
  // その他の設定
  console.log('Session Timeout:', properties.SESSION_TIMEOUT_MINUTES + '分');
  console.log('Max Participants:', properties.MAX_PARTICIPANTS_PER_SESSION + '人');
  console.log('App Version:', properties.APP_VERSION);
  console.log('Debug Mode:', properties.DEBUG_MODE);
  
  return properties;
}

/**
 * 設定値のバリデーション
 * 必要な設定が正しく行われているかチェック
 */
function validateConfiguration() {
  const properties = PropertiesService.getScriptProperties().getProperties();
  const validationResults = {
    isValid: true,
    errors: [],
    warnings: []
  };
  
  // Gemini API Key チェック
  const apiKey = properties.GEMINI_API_KEY;
  if (!apiKey || apiKey === 'YOUR_GEMINI_API_KEY_HERE') {
    validationResults.isValid = false;
    validationResults.errors.push('Gemini API Keyが設定されていません');
  } else if (apiKey.length < 30) {
    validationResults.warnings.push('API Keyの形式が正しくない可能性があります');
  }
  
  // スプレッドシートID チェック
  const spreadsheetId = properties.SPREADSHEET_ID;
  if (!spreadsheetId || spreadsheetId === 'YOUR_SPREADSHEET_ID_HERE') {
    validationResults.isValid = false;
    validationResults.errors.push('スプレッドシートIDが設定されていません');
  } else {
    try {
      SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      validationResults.isValid = false;
      validationResults.errors.push('スプレッドシートにアクセスできません: ' + error.message);
    }
  }
  
  // 数値設定のチェック
  const timeout = parseInt(properties.SESSION_TIMEOUT_MINUTES);
  if (isNaN(timeout) || timeout < 1 || timeout > 30) {
    validationResults.warnings.push('セッションタイムアウト値が推奨範囲外です（1-30分）');
  }
  
  const maxParticipants = parseInt(properties.MAX_PARTICIPANTS_PER_SESSION);
  if (isNaN(maxParticipants) || maxParticipants < 1 || maxParticipants > 100) {
    validationResults.warnings.push('最大参加者数が推奨範囲外です（1-100人）');
  }
  
  // 結果表示
  if (validationResults.isValid) {
    console.log('✅ 設定値の検証が完了しました（問題なし）');
  } else {
    console.log('❌ 設定に問題があります:');
    validationResults.errors.forEach(error => console.log('  - ' + error));
  }
  
  if (validationResults.warnings.length > 0) {
    console.log('⚠️  警告:');
    validationResults.warnings.forEach(warning => console.log('  - ' + warning));
  }
  
  return validationResults;
}

/**
 * 設定状況の詳細チェック
 * 各プロパティの設定状況を確認し、不足項目を報告
 */
function checkPropertiesStatus() {
  const properties = PropertiesService.getScriptProperties().getProperties();
  const status = {
    configured: [],
    missing: [],
    warnings: []
  };
  
  console.log('=== スクリプトプロパティ設定状況 ===\n');
  
  Object.entries(REQUIRED_PROPERTIES).forEach(([key, config]) => {
    const value = properties[key];
    
    if (value) {
      status.configured.push(key);
      if (config.sensitive) {
        // 機密情報は一部のみ表示
        const displayValue = value.length > 8 ? 
          `${value.substring(0, 4)}...${value.substring(value.length - 4)}` : 
          '***';
        console.log(`✅ ${key}: ${displayValue}`);
      } else {
        console.log(`✅ ${key}: ${value}`);
      }
    } else {
      if (config.required) {
        status.missing.push(key);
        console.log(`❌ ${key}: 未設定 [必須]`);
      } else {
        status.warnings.push(key);
        console.log(`⚠️  ${key}: 未設定 (デフォルト値: ${config.defaultValue || 'なし'})`);
      }
    }
  });
  
  console.log('\n=== 設定状況サマリー ===');
  console.log(`設定済み: ${status.configured.length}項目`);
  console.log(`未設定（必須）: ${status.missing.length}項目`);
  console.log(`未設定（オプション）: ${status.warnings.length}項目`);
  
  if (status.missing.length > 0) {
    console.log('\n⚠️  以下の必須項目をGAS GUIから設定してください:');
    status.missing.forEach(key => {
      console.log(`  - ${key}: ${REQUIRED_PROPERTIES[key].description}`);
    });
  }
  
  return status;
}

/**
 * 設定リセット（開発・テスト用）
 * 全ての設定値をデフォルトに戻す
 */
function resetConfiguration() {
  try {
    PropertiesService.getScriptProperties().deleteAll();
    configureScriptProperties();
    console.log('✅ 設定値をリセットしました');
    return { success: true, message: '設定値を初期状態にリセットしました' };
  } catch (error) {
    console.error('❌ 設定リセットエラー:', error);
    return { success: false, message: '設定リセット中にエラーが発生しました: ' + error.toString() };
  }
}