/**
 * Commands JavaScript
 * リボンボタンからのコマンド実行を処理
 */

import { ConfigService } from '../services/ConfigService.js';
import { OutlookEventService } from '../services/OutlookEventService.js';
import { SharePointService } from '../services/SharePointService.js';
import { PowerAutomateService } from '../services/PowerAutomateService.js';
import { LoggingService } from '../services/LoggingService.js';

// グローバルなサービスインスタンス
let services = {
    config: null,
    outlook: null,
    sharePoint: null,
    powerAutomate: null,
    logger: null
};

/**
 * Office.js の準備完了時の処理
 */
Office.onReady((info) => {
    console.log('Commands: Office.js is ready');
    
    // サービスを初期化
    initializeServices().catch(error => {
        console.error('Commands: Service initialization failed', error);
    });
});

/**
 * サービスを初期化
 */
async function initializeServices() {
    try {
        // ロギングサービス
        services.logger = LoggingService.getInstance();
        services.logger.info('Commands: Initializing services...');
        
        // 設定サービス
        services.config = ConfigService.getInstance();
        await services.config.initialize();
        
        // Outlook サービス
        services.outlook = new OutlookEventService();
        await services.outlook.initialize();
        
        // SharePoint サービス
        services.sharePoint = new SharePointService();
        await services.sharePoint.initialize();
        
        // Power Automate サービス
        services.powerAutomate = new PowerAutomateService();
        await services.powerAutomate.initialize();
        
        services.logger.info('Commands: All services initialized successfully');
        
    } catch (error) {
        console.error('Commands: Failed to initialize services', error);
        if (services.logger) {
            services.logger.error('Commands service initialization failed', error);
        }
    }
}

/**
 * クイック処理コマンド - 現在の会議の外部ユーザーを即座に処理
 */
async function quickProcessVisitors(event) {
    try {
        services.logger?.info('Commands: Quick process visitors started');
        
        // サービスが初期化されているかチェック
        if (!services.outlook || !services.sharePoint || !services.powerAutomate) {
            throw new Error('サービスが初期化されていません');
        }
        
        // 現在の会議データを取得
        const meetingData = await services.outlook.getCurrentEventData();
        
        if (!meetingData) {
            showNotification('warning', '会議データが見つかりません', '現在選択されている会議を確認してください。');
            return;
        }
        
        if (!meetingData.externalUsers || meetingData.externalUsers.length === 0) {
            showNotification('info', '外部ユーザーなし', 'この会議には外部ユーザーが含まれていません。');
            return;
        }
        
        services.logger?.info(`Commands: Processing ${meetingData.externalUsers.length} external users`);
        
        // SharePoint にレコードを作成/更新
        const sharePointResult = await services.sharePoint.updateVisitorRecords(
            meetingData.meetingId,
            meetingData.subject,
            meetingData.externalUsers,
            meetingData.startTime,
            meetingData.endTime
        );
        
        // Power Automate で通知を送信
        const notificationResult = await services.powerAutomate.sendVisitorNotification(
            meetingData.meetingId,
            meetingData.subject,
            meetingData.externalUsers,
            meetingData.startTime,
            meetingData.endTime,
            'created'
        );
        
        services.logger?.info(`Commands: Quick process completed - SharePoint: ${sharePointResult?.length || 0}, Notifications: ${notificationResult.notificationsSent}`);
        
        // 成功通知
        showNotification('success', '処理完了', 
            `${meetingData.externalUsers.length}人の外部ユーザーが登録され、通知が送信されました。`);
            
    } catch (error) {
        services.logger?.error('Commands: Quick process failed', error);
        showNotification('error', '処理エラー', `エラーが発生しました: ${error.message}`);
    } finally {
        // コマンドの完了を通知
        if (event) {
            event.completed();
        }
    }
}

/**
 * 設定テストコマンド - SharePoint と Power Automate の接続をテスト
 */
async function testConnections(event) {
    try {
        services.logger?.info('Commands: Connection test started');
        
        const results = [];
        
        // SharePoint接続テスト
        try {
            if (services.sharePoint) {
                const sharePointTest = await services.sharePoint.testConnection();
                results.push({
                    service: 'SharePoint',
                    success: sharePointTest.success,
                    message: sharePointTest.success 
                        ? `接続成功 (${sharePointTest.itemCount}件のアイテム)` 
                        : `接続失敗: ${sharePointTest.error}`
                });
            } else {
                results.push({
                    service: 'SharePoint',
                    success: false,
                    message: 'サービスが初期化されていません'
                });
            }
        } catch (error) {
            results.push({
                service: 'SharePoint',
                success: false,
                message: `テストエラー: ${error.message}`
            });
        }
        
        // Power Automate接続テスト
        try {
            if (services.powerAutomate) {
                const powerAutomateTest = await services.powerAutomate.testConnection();
                results.push({
                    service: 'Power Automate',
                    success: powerAutomateTest.success,
                    message: powerAutomateTest.success 
                        ? `接続成功 (ステータス: ${powerAutomateTest.statusCode})` 
                        : `接続失敗: ${powerAutomateTest.error}`
                });
            } else {
                results.push({
                    service: 'Power Automate',
                    success: false,
                    message: 'サービスが初期化されていません'
                });
            }
        } catch (error) {
            results.push({
                service: 'Power Automate',
                success: false,
                message: `テストエラー: ${error.message}`
            });
        }
        
        services.logger?.info('Commands: Connection test completed', results);
        
        // 結果を表示
        const allSuccess = results.every(r => r.success);
        const summaryMessage = results.map(r => 
            `${r.service}: ${r.success ? '✅' : '❌'} ${r.message}`
        ).join('\n');
        
        showNotification(
            allSuccess ? 'success' : 'warning', 
            '接続テスト結果',
            summaryMessage
        );
        
    } catch (error) {
        services.logger?.error('Commands: Connection test failed', error);
        showNotification('error', 'テストエラー', `接続テストでエラーが発生しました: ${error.message}`);
    } finally {
        // コマンドの完了を通知
        if (event) {
            event.completed();
        }
    }
}

/**
 * 統計表示コマンド - SharePoint リストの統計情報を表示
 */
async function showStatistics(event) {
    try {
        services.logger?.info('Commands: Show statistics started');
        
        if (!services.sharePoint) {
            throw new Error('SharePointサービスが初期化されていません');
        }
        
        const stats = await services.sharePoint.getStatistics(7);
        
        const message = `過去7日間の統計:
📊 総来訪者: ${stats.total}人
📅 予定済み: ${stats.scheduled}人  
✅ 完了済み: ${stats.completed}人
❌ キャンセル: ${stats.cancelled}人`;
        
        services.logger?.info('Commands: Statistics retrieved', stats);
        
        showNotification('info', '統計情報', message);
        
    } catch (error) {
        services.logger?.error('Commands: Show statistics failed', error);
        showNotification('error', '統計エラー', `統計情報の取得でエラーが発生しました: ${error.message}`);
    } finally {
        // コマンドの完了を通知
        if (event) {
            event.completed();
        }
    }
}

/**
 * 設定表示コマンド - 現在の設定情報を表示
 */
async function showConfiguration(event) {
    try {
        services.logger?.info('Commands: Show configuration started');
        
        if (!services.config) {
            throw new Error('設定サービスが初期化されていません');
        }
        
        const config = services.config.getConfig();
        
        const message = `現在の設定:
🌐 SharePoint: ${config.sharePoint?.siteUrl ? '設定済み' : '未設定'}
📝 リスト名: ${config.sharePoint?.listName || '未設定'}
🔄 Power Automate: ${config.powerAutomate?.notificationUrl ? '設定済み' : '未設定'}
🏢 内部ドメイン: ${config.outlook?.internalDomains?.length || 0}個
📋 ログレベル: ${config.logging?.level || '未設定'}`;
        
        services.logger?.info('Commands: Configuration displayed');
        
        showNotification('info', '設定情報', message);
        
    } catch (error) {
        services.logger?.error('Commands: Show configuration failed', error);
        showNotification('error', '設定エラー', `設定情報の表示でエラーが発生しました: ${error.message}`);
    } finally {
        // コマンドの完了を通知
        if (event) {
            event.completed();
        }
    }
}

/**
 * 手動同期コマンド - 現在の会議を強制的に同期
 */
async function manualSync(event) {
    try {
        services.logger?.info('Commands: Manual sync started');
        
        if (!services.outlook || !services.sharePoint) {
            throw new Error('必要なサービスが初期化されていません');
        }
        
        // 現在の会議データを取得
        const meetingData = await services.outlook.getCurrentEventData();
        
        if (!meetingData) {
            showNotification('warning', '同期不可', '同期する会議データが見つかりません。');
            return;
        }
        
        // 既存のレコードを取得
        const existingRecords = await services.sharePoint.getVisitorRecordsByMeetingId(meetingData.meetingId);
        
        let message;
        if (meetingData.externalUsers && meetingData.externalUsers.length > 0) {
            // 外部ユーザーがいる場合は更新
            await services.sharePoint.updateVisitorRecords(
                meetingData.meetingId,
                meetingData.subject,
                meetingData.externalUsers,
                meetingData.startTime,
                meetingData.endTime
            );
            
            message = `同期完了: ${meetingData.externalUsers.length}人の外部ユーザーを同期しました。`;
        } else {
            // 外部ユーザーがいない場合は削除
            if (existingRecords.length > 0) {
                await services.sharePoint.deleteVisitorRecords(meetingData.meetingId);
                message = `同期完了: ${existingRecords.length}件のレコードを削除しました（外部ユーザーなし）。`;
            } else {
                message = '同期完了: 変更はありませんでした。';
            }
        }
        
        services.logger?.info('Commands: Manual sync completed');
        
        showNotification('success', '同期完了', message);
        
    } catch (error) {
        services.logger?.error('Commands: Manual sync failed', error);
        showNotification('error', '同期エラー', `同期処理でエラーが発生しました: ${error.message}`);
    } finally {
        // コマンドの完了を通知
        if (event) {
            event.completed();
        }
    }
}

/**
 * 通知を表示
 */
function showNotification(type, title, message) {
    try {
        // Office.js の通知機能を使用
        if (Office.context.ui && Office.context.ui.displayDialogAsync) {
            // ダイアログで表示（詳細メッセージ用）
            const dialogHtml = `
                <!DOCTYPE html>
                <html>
                <head>
                    <title>${title}</title>
                    <style>
                        body { 
                            font-family: 'Segoe UI', sans-serif; 
                            padding: 20px; 
                            background-color: ${type === 'success' ? '#dff6dd' : type === 'error' ? '#fde7e9' : type === 'warning' ? '#fff4ce' : '#f3f2f1'};
                        }
                        h2 { 
                            color: ${type === 'success' ? '#107c10' : type === 'error' ? '#a4262c' : type === 'warning' ? '#8a8886' : '#323130'};
                            margin-top: 0;
                        }
                        pre { 
                            background: white; 
                            padding: 10px; 
                            border-radius: 4px; 
                            white-space: pre-wrap; 
                            border: 1px solid #edebe9;
                        }
                        .close-btn {
                            background-color: #0078d4;
                            color: white;
                            border: none;
                            padding: 8px 16px;
                            border-radius: 4px;
                            cursor: pointer;
                            margin-top: 15px;
                        }
                    </style>
                </head>
                <body>
                    <h2>${title}</h2>
                    <pre>${message}</pre>
                    <button class="close-btn" onclick="window.close()">閉じる</button>
                </body>
                </html>
            `;
            
            const blob = new Blob([dialogHtml], { type: 'text/html' });
            const url = URL.createObjectURL(blob);
            
            Office.context.ui.displayDialogAsync(url, 
                { height: 300, width: 400 },
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        setTimeout(() => {
                            URL.revokeObjectURL(url);
                        }, 1000);
                    }
                }
            );
            
        } else if (Office.context.mailbox && Office.context.mailbox.item && Office.context.mailbox.item.notificationMessages) {
            // 通知メッセージを使用（簡潔なメッセージ用）
            const notificationId = `lobby_${Date.now()}`;
            const shortMessage = message.length > 150 ? message.substring(0, 147) + '...' : message;
            
            Office.context.mailbox.item.notificationMessages.addAsync(notificationId, {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: `${title}: ${shortMessage}`,
                icon: type === 'success' ? 'icon1' : type === 'error' ? 'icon2' : 'icon3',
                persistent: false
            });
            
            // 5秒後に削除
            setTimeout(() => {
                Office.context.mailbox.item.notificationMessages.removeAsync(notificationId);
            }, 5000);
        } else {
            // フォールバック: コンソールに出力
            console.log(`${type.toUpperCase()}: ${title} - ${message}`);
        }
        
    } catch (error) {
        console.error('Failed to show notification:', error);
        console.log(`${type.toUpperCase()}: ${title} - ${message}`);
    }
}

/**
 * エラーハンドリング用のグローバル関数
 */
function handleCommandError(commandName, error, event) {
    const errorMessage = `コマンド "${commandName}" でエラーが発生しました: ${error.message}`;
    
    services.logger?.error(`Commands: ${commandName} failed`, error);
    console.error(errorMessage, error);
    
    showNotification('error', 'コマンドエラー', errorMessage);
    
    if (event) {
        event.completed();
    }
}

// グローバル関数として登録（manifest.xml から呼び出し可能にする）
window.quickProcessVisitors = (event) => quickProcessVisitors(event).catch(error => handleCommandError('quickProcessVisitors', error, event));
window.testConnections = (event) => testConnections(event).catch(error => handleCommandError('testConnections', error, event));
window.showStatistics = (event) => showStatistics(event).catch(error => handleCommandError('showStatistics', error, event));
window.showConfiguration = (event) => showConfiguration(event).catch(error => handleCommandError('showConfiguration', error, event));
window.manualSync = (event) => manualSync(event).catch(error => handleCommandError('manualSync', error, event));

// デバッグ用
window.lobbyCommands = {
    services,
    quickProcessVisitors,
    testConnections,
    showStatistics,
    showConfiguration,
    manualSync
};