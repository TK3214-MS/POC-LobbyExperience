# Lobby Experience - Outlook来訪者管理アドイン

## 概要

Lobby Experience は、Microsoft Outlook と連携して外部来訪者の情報を自動的に管理するアドインです。Outlook で会議を作成・編集すると、外部参加者の情報を SharePoint リストに自動登録し、Power Automate 経由で関係者に通知を送信します。

## 主な機能

- ✅ **自動外部ユーザー検出**: 会議の参加者から内部ドメイン以外のユーザーを自動抽出
- ✅ **SharePoint 連携**: 来訪者情報を SharePoint リストに自動登録・更新・削除
- ✅ **Power Automate 通知**: 外部ユーザーへの自動通知送信
- ✅ **リアルタイム監視**: 会議内容の変更を自動検知して同期
- ✅ **リボンコマンド**: ワンクリックでの手動処理・テスト機能
- ✅ **統計・ログ機能**: 処理状況の可視化とトラブルシューティング支援

## システム要件

### 必要なソフトウェア
- Microsoft Outlook (Microsoft 365 または Outlook 2019/2021)
- Node.js 18.0 以降
- npm または yarn

### 必要な Microsoft 365 サービス
- SharePoint Online (来訪者情報リスト用)
- Power Automate (通知送信用)
- Azure Active Directory (認証用)

## インストールと設定

### 1. プロジェクトのセットアップ

```bash
# リポジトリをクローン
git clone <repository-url>
cd POC-LobbyExperience

# 依存関係をインストール
npm install

# 開発用証明書を生成
npm run dev-certs
```

### 2. SharePoint リスト設定

SharePoint サイトに以下の名前とカラムを持つリストを作成してください：

#### リスト名: `LobbyVisitors`

| カラム名 | データ型 | 説明 | 必須 |
|---------|---------|------|------|
| MeetingId | 1行テキスト | 会議の一意識別子 | ✅ |
| MeetingTitle | 1行テキスト | 会議のタイトル | ✅ |
| VisitorEmail | 1行テキスト | 来訪者のメールアドレス | ✅ |
| VisitorName | 1行テキスト | 来訪者の表示名 | ❌ |
| StartTime | 日付と時刻 | 会議開始時刻 | ✅ |
| EndTime | 日付と時刻 | 会議終了時刻 | ✅ |
| Status | 選択肢 | 来訪状況 (Scheduled, Completed, Cancelled) | ❌ |
| CreatedDate | 日付と時刻 | レコード作成日時 | ❌ |
| ModifiedDate | 日付と時刻 | レコード更新日時 | ❌ |

#### SharePoint リストの作成手順

1. SharePoint サイトにアクセス
2. 「新規」→「リスト」をクリック
3. 「空白のリスト」を選択
4. リスト名を「`LobbyVisitors`」に設定
5. 上記のカラムを順次追加

#### カラム作成のサンプル PowerShell スクリプト

```powershell
# SharePoint PnP PowerShell を使用したリスト作成例
Connect-PnPOnline -Url "https://yourcompany.sharepoint.com/sites/yoursite"

# リスト作成
New-PnPList -Title "LobbyVisitors" -Template GenericList

# カラム追加
Add-PnPField -List "LobbyVisitors" -DisplayName "MeetingId" -InternalName "MeetingId" -Type Text -Required
Add-PnPField -List "LobbyVisitors" -DisplayName "MeetingTitle" -InternalName "MeetingTitle" -Type Text -Required
Add-PnPField -List "LobbyVisitors" -DisplayName "VisitorEmail" -InternalName "VisitorEmail" -Type Text -Required
Add-PnPField -List "LobbyVisitors" -DisplayName "VisitorName" -InternalName "VisitorName" -Type Text
Add-PnPField -List "LobbyVisitors" -DisplayName "StartTime" -InternalName "StartTime" -Type DateTime -Required
Add-PnPField -List "LobbyVisitors" -DisplayName "EndTime" -InternalName "EndTime" -Type DateTime -Required
Add-PnPField -List "LobbyVisitors" -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Scheduled","Completed","Cancelled"
Add-PnPField -List "LobbyVisitors" -DisplayName "CreatedDate" -InternalName "CreatedDate" -Type DateTime
Add-PnPField -List "LobbyVisitors" -DisplayName "ModifiedDate" -InternalName "ModifiedDate" -Type DateTime
```

### 3. Power Automate フロー設定

外部ユーザーへの通知用 Power Automate フローを作成してください：

#### トリガー設定
- **トリガータイプ**: HTTP要求の受信時
- **要求本文のJSONスキーマ**:

```json
{
    "type": "object",
    "properties": {
        "meetingId": {"type": "string"},
        "meetingTitle": {"type": "string"},
        "visitorEmail": {"type": "string"},
        "visitorName": {"type": "string"},
        "startTime": {"type": "string"},
        "endTime": {"type": "string"},
        "notificationType": {"type": "string"},
        "timestamp": {"type": "string"},
        "source": {"type": "string"}
    }
}
```

#### アクション例
1. **Outlook でメール送信** - 来訪者への通知メール
2. **Teams メッセージ送信** - 受付担当者への通知
3. **SharePoint リスト更新** - 通知状況の記録

#### サンプル通知メールテンプレート

```html
件名: 【来訪予定】@{triggerBody()?['meetingTitle']} のお知らせ

@{triggerBody()?['visitorName']} 様

下記の会議にご参加いただく予定です。

■ 会議情報
・件名: @{triggerBody()?['meetingTitle']}
・開始時刻: @{triggerBody()?['startTime']}
・終了時刻: @{triggerBody()?['endTime']}

ご来社の際は、1階受付にお声かけください。

よろしくお願いいたします。
```

### 4. Azure AD アプリケーション登録

SharePoint API アクセス用の Azure AD アプリを登録してください：

1. Azure Portal → Azure Active Directory → アプリの登録
2. 「新規登録」をクリック
3. アプリケーション名を設定（例：LobbyExperience）
4. リダイレクト URI を設定：`https://localhost:3000` (開発時)
5. API のアクセス許可を追加：
   - Microsoft Graph: `Sites.ReadWrite.All`
   - SharePoint: `Sites.ReadWrite.All`

### 5. 設定ファイルの作成

`config.sample.json` をコピーして `config.json` を作成し、環境に合わせて設定してください：

```json
{
  "sharePoint": {
    "siteUrl": "https://yourcompany.sharepoint.com/sites/yoursite",
    "listName": "LobbyVisitors",
    "clientId": "your-azure-app-client-id",
    "tenantId": "your-tenant-id"
  },
  "powerAutomate": {
    "notificationUrl": "https://prod-xx.eastus.logic.azure.com:443/workflows/xxxxxxxx/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=xxxxxxxxx",
    "timeout": 30000
  },
  "outlook": {
    "internalDomains": [
      "yourcompany.com",
      "subsidiary.com"
    ],
    "excludeRooms": true
  },
  "logging": {
    "level": "info",
    "enableConsole": true
  }
}
```

#### 設定項目の詳細

| 設定項目 | 説明 | 例 |
|---------|------|---|
| `sharePoint.siteUrl` | SharePoint サイトの URL | `https://contoso.sharepoint.com/sites/lobby` |
| `sharePoint.listName` | 来訪者情報リストの名前 | `LobbyVisitors` |
| `sharePoint.clientId` | Azure AD アプリの Client ID | `12345678-1234-1234-1234-123456789abc` |
| `sharePoint.tenantId` | Azure AD テナント ID | `87654321-4321-4321-4321-cba987654321` |
| `powerAutomate.notificationUrl` | Power Automate フローの HTTP トリガー URL | `https://prod-xx.eastus.logic.azure.com:443/workflows/...` |
| `powerAutomate.timeout` | API 呼び出しタイムアウト (ミリ秒) | `30000` |
| `outlook.internalDomains` | 内部ドメインのリスト（これ以外を外部ユーザーとして扱う） | `["contoso.com", "subsidiary.com"]` |
| `outlook.excludeRooms` | 会議室を外部ユーザーから除外するか | `true` |
| `logging.level` | ログレベル (`debug`, `info`, `warn`, `error`) | `info` |
| `logging.enableConsole` | コンソール出力を有効にするか | `true` |

## 開発とテスト

### 開発環境の起動

```bash
# 開発サーバーを起動
npm run dev-server

# Outlook でアドインを起動（別ターミナル）
npm start
```

### ビルドとデプロイ

```bash
# 本番用ビルド
npm run build

# マニフェストファイルの検証
npm run validate

# アドインの停止
npm stop
```

### テスト方法

#### 1. 単体テスト
```bash
# Jest を使用した単体テスト実行
npm test
```

#### 2. 手動テスト手順

**基本動作のテスト：**

1. Outlook で新しい会議を作成
2. 件名と時間を設定
3. 内部ユーザーと外部ユーザーを参加者に追加
4. アドインのタスクペーンを開く
5. 「来訪者を処理」ボタンをクリック
6. SharePoint リストに外部ユーザーのレコードが作成されることを確認
7. Power Automate で通知が送信されることを確認

**変更検知のテスト：**

1. 既存の会議の参加者を変更（外部ユーザーの追加・削除）
2. 件名や時間を変更
3. 自動的に SharePoint の情報が同期されることを確認

**リボンコマンドのテスト：**

1. 会議を選択した状態で「ホーム」リボンの「Lobby Experience」セクションを確認
2. 各ボタンをクリックして正常動作することを確認：
   - 「クイック処理」- 即座に来訪者情報を処理
   - 「接続テスト」- SharePoint/Power Automate への接続確認
   - 「統計表示」- 処理実績の表示
   - 「設定表示」- 現在の設定内容の表示
   - 「手動同期」- 強制的な情報同期

#### 3. 接続テスト

```bash
# PowerShell での接続テスト例
$config = Get-Content "config.json" | ConvertFrom-Json
$headers = @{
    "Content-Type" = "application/json"
    "Accept" = "application/json"
}
$body = @{
    test = $true
    timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    source = "Manual Test"
} | ConvertTo-Json

Invoke-RestMethod -Uri $config.powerAutomate.notificationUrl -Method Post -Headers $headers -Body $body
```

## トラブルシューティング

### よくある問題と解決方法

#### 1. アドインが読み込まれない

**原因:**
- manifest.xml の設定ミス
- 開発証明書の問題
- ポート番号の競合

**解決方法:**
```bash
# マニフェストの検証
npm run validate

# 証明書の再生成
npm run dev-certs

# ポート番号の確認
netstat -an | findstr :3000
```

#### 2. SharePoint に接続できない

**原因:**
- Azure AD アプリの権限不足
- config.json の設定ミス
- ネットワークの制限

**解決方法:**
1. Azure Portal でアプリの API 権限を確認
2. 管理者による同意が必要な場合は管理者に依頼
3. config.json の siteUrl と listName を再確認
4. ブラウザで直接 SharePoint サイトにアクセスできるか確認

#### 3. Power Automate 通知が送信されない

**原因:**
- フローが無効化されている
- HTTP トリガー URL が間違っている
- JSON スキーマの不整合

**解決方法:**
1. Power Automate でフローの実行履歴を確認
2. HTTP トリガー URL を再取得して config.json を更新
3. フローのテスト実行で JSON 形式を確認

#### 4. 外部ユーザーが正しく検出されない

**原因:**
- internalDomains の設定ミス
- 会議室が外部ユーザーとして判定されている

**解決方法:**
1. config.json の internalDomains に所属ドメインをすべて追加
2. excludeRooms を true に設定
3. アドインのログで検出されているユーザーを確認

### ログの確認方法

#### 1. ブラウザーコンソール
1. F12 キーでデベロッパーツールを開く
2. Console タブでエラーメッセージを確認

#### 2. アドイン内ログ
1. タスクペーンの「活動ログ」セクションを確認
2. 処理の流れとエラー内容を把握

#### 3. Power Automate 実行履歴
1. Power Automate ポータルにアクセス
2. 該当フローの実行履歴を確認
3. 失敗した場合はエラーの詳細を確認

## セキュリティ考慮事項

### 1. 認証とアクセス制御
- Azure AD によるシングルサインオン
- 最小権限の原則に基づく API 権限設定
- アクセストークンの適切な管理と更新

### 2. データ保護
- HTTPS 通信の強制
- 個人情報の最小限収集
- データの暗号化転送

### 3. 監査とログ
- すべての操作をログ記録
- 異常なアクセスパターンの検知
- 定期的なアクセス権限の見直し

## パフォーマンス最適化

### 1. API 呼び出しの最適化
- バッチ処理での一括更新
- キャッシュ機能の活用
- リトライ機能によるエラーハンドリング

### 2. ユーザー体験の向上
- 非同期処理によるブロッキング回避
- プログレス表示の実装
- エラーメッセージの分かりやすい表示

## 更新とメンテナンス

### 1. 定期メンテナンス項目
- 依存関係のセキュリティ更新
- SharePoint リストのパフォーマンス確認
- ログファイルのクリーンアップ

### 2. 更新手順
```bash
# 依存関係の更新
npm update

# セキュリティ監査
npm audit

# 脆弱性の修正
npm audit fix
```

## サポートと問い合わせ

### 技術サポート
- **GitHub Issues**: バグ報告や機能要望
- **Wiki**: 詳細なドキュメントと FAQ
- **メール**: <support@yourcompany.com>

### コントリビューション
1. Fork this repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## ライセンス

このプロジェクトは MIT ライセンスの下で公開されています。詳細は [LICENSE](LICENSE) ファイルを参照してください。

## バージョン履歴

- **v1.0.0** (2025-09-09): 初回リリース
  - 基本的な来訪者管理機能
  - SharePoint 連携
  - Power Automate 通知
  - リアルタイム変更検知

---

**注意**: このアドインを本番環境で使用する前に、必ずテスト環境で十分な検証を行ってください。