# テストデータとサンプル

このディレクトリには、Lobby Experience アドインの開発とテストに使用するサンプルデータファイルが含まれています。

## ファイル一覧

### 設定ファイル

#### `test-config.json`
テスト環境用の設定ファイル。本番環境とは別の SharePoint サイトや Power Automate フローを使用してテストを行うためのサンプル設定です。

**使用方法:**
```bash
# テスト設定を使用してアドインを起動
cp tests/test-config.json config.json
npm start
```

### テストデータファイル

#### `sample-meeting-data.json`
Outlook 会議の典型的なデータ構造のサンプル。外部ユーザーと内部ユーザー、会議室リソースが含まれています。

**主要項目:**
- 会議基本情報（ID、件名、時刻）
- 全参加者リスト
- 外部ユーザーのフィルタリング結果
- 主催者情報
- 場所・オンライン会議設定

#### `sharepoint-test-data.json`
SharePoint リストのサンプルデータと構造定義。

**内容:**
- 来訪者レコードのサンプル（4件）
- SharePoint リストのフィールド定義
- よく使用される OData クエリサンプル

**サンプルレコード状態:**
- Scheduled（予定済み）: 新規登録された来訪者
- Completed（完了）: 来訪完了した訪問者  
- Cancelled（キャンセル）: キャンセルされた予定

#### `power-automate-test-payload.json`
Power Automate フローに送信するペイロードのサンプル。

**通知タイプ:**
- `created`: 新規来訪者の通知
- `updated`: 会議情報変更の通知
- `cancelled`: 会議キャンセルの通知

**テンプレート:**
- メール件名・本文のテンプレート
- Teams メッセージのテンプレート
- 各種シナリオのテストケース

#### `mock-office-data.json`
Office.js API の模擬データ。単体テストや開発時のモック作成に使用。

**モックデータ:**
- Office Context の構造
- SharePoint API レスポンス
- Power Automate API レスポンス
- テストユーザー（内部・外部・会議室）
- イベントデータのサンプル

## テストシナリオ

### 1. 基本フロー テスト

```javascript
// サンプル会議データを使用した基本テスト
const meetingData = require('./sample-meeting-data.json');

// 外部ユーザーが正しく検出されるかテスト
console.log('External users:', meetingData.externalUsers);
// 出力: John Smith, Jane Doe (会議室は除外)
```

### 2. SharePoint 連携テスト

```powershell
# PowerShell を使用した SharePoint API テスト
$config = Get-Content "test-config.json" | ConvertFrom-Json
$siteUrl = $config.sharePoint.siteUrl
$listName = $config.sharePoint.listName

# リスト構造の確認
Invoke-RestMethod -Uri "$siteUrl/_api/web/lists/getByTitle('$listName')/fields" -Headers @{Authorization="Bearer $token"}
```

### 3. Power Automate 通知テスト

```powershell
# HTTP トリガーテスト
$testPayload = Get-Content "power-automate-test-payload.json" | ConvertFrom-Json
$notificationUrl = $config.powerAutomate.notificationUrl

$headers = @{
    "Content-Type" = "application/json"
    "Accept" = "application/json"
}

# 新規通知テスト
$body = $testPayload.testPayloads.newVisitorNotification | ConvertTo-Json
Invoke-RestMethod -Uri $notificationUrl -Method Post -Headers $headers -Body $body
```

### 4. 統合テスト

```bash
# Jest を使用した自動テスト
npm test

# 特定のテストファイルのみ実行
npm test -- --testPathPattern=SharePointService.test.js
```

## モックデータの使用

### Office.js モック

```javascript
// テスト用の Office.js モック設定
const mockOfficeData = require('./mock-office-data.json');

// Office Context のモック
global.Office = {
    context: mockOfficeData.mockOfficeContext,
    onReady: (callback) => callback({ host: 'Outlook', platform: 'PC' })
};
```

### API レスポンスのモック

```javascript
// fetch のモック設定
global.fetch = jest.fn((url) => {
    if (url.includes('sharepoint.com')) {
        return Promise.resolve({
            ok: true,
            json: () => Promise.resolve(mockOfficeData.mockApiResponses.sharePoint.listItems)
        });
    }
    
    if (url.includes('logic.azure.com')) {
        return Promise.resolve({
            ok: true,
            json: () => Promise.resolve(mockOfficeData.mockApiResponses.powerAutomate.success)
        });
    }
});
```

## データ生成ツール

### 1. ランダムテストデータ生成

```javascript
// 新しいテスト会議データを生成
function generateTestMeeting() {
    const meetingId = `meeting_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    const startTime = new Date(Date.now() + 24 * 60 * 60 * 1000); // 明日
    const endTime = new Date(startTime.getTime() + 2 * 60 * 60 * 1000); // 2時間後
    
    return {
        meetingId,
        subject: `テスト会議_${meetingId}`,
        startTime: startTime.toISOString(),
        endTime: endTime.toISOString(),
        externalUsers: [
            {
                emailAddress: `test${Math.floor(Math.random() * 1000)}@external.com`,
                name: `Test User ${Math.floor(Math.random() * 100)}`
            }
        ]
    };
}
```

### 2. SharePoint テストデータ投入

```javascript
// テストデータを SharePoint に投入
async function seedTestData() {
    const sharePointService = new SharePointService();
    const testData = require('./sharepoint-test-data.json');
    
    for (const record of testData.sampleVisitorRecords) {
        await sharePointService.createListItem(record);
    }
    
    console.log('Test data seeded successfully');
}
```

## 注意事項

### セキュリティ

- テストファイルには実際の認証情報を含めないでください
- 本番環境の URL や機密情報はサンプルデータから除外してください
- テスト用の Azure AD アプリケーションを使用してください

### データ管理

- テスト実行後はテストデータをクリーンアップしてください
- 本番データとテストデータを混同しないよう注意してください
- 定期的にテストデータの整合性を確認してください

### パフォーマンス

- 大量のテストデータを使用する際はバッチ処理を検討してください
- モックデータを活用してネットワーク呼び出しを最小限に抑えてください
- テスト実行時間を短縮するため、必要最小限のデータのみ使用してください

## 問い合わせ

テストデータの追加や修正については、プロジェクトの Issue で報告してください。