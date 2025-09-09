/**
 * Taskpane JavaScript
 * ユーザーインターフェースのメイン処理
 */

import { ConfigService } from '../services/ConfigService.js';
import { OutlookEventService } from '../services/OutlookEventService.js';
import { SharePointService } from '../services/SharePointService.js';
import { PowerAutomateService } from '../services/PowerAutomateService.js';
import { LoggingService } from '../services/LoggingService.js';

class TaskpaneApp {
    constructor() {
        this.configService = null;
        this.outlookService = null;
        this.sharePointService = null;
        this.powerAutomateService = null;
        this.logger = null;
        
        this.currentMeetingData = null;
        this.isProcessing = false;
        
        this.elements = {};
    }

    /**
     * アプリケーションを初期化
     */
    async initialize() {
        try {
            // Office.js の準備完了を待つ
            await new Promise((resolve) => {
                Office.onReady((info) => {
                    resolve(info);
                });
            });

            // UI要素を取得
            this.initializeElements();
            
            // イベントリスナーを設定
            this.setupEventListeners();
            
            // ロギングサービスを初期化
            this.logger = LoggingService.getInstance();
            this.setupLogHandler();
            
            this.addLog('info', 'アプリケーションを初期化しています...');
            
            // サービスを初期化
            await this.initializeServices();
            
            // UI を更新
            await this.refreshMeetingInfo();
            await this.refreshServiceStatus();
            await this.refreshSettings();
            
            this.addLog('info', 'アプリケーションの初期化が完了しました');
            
        } catch (error) {
            this.addLog('error', `初期化エラー: ${error.message}`);
            this.showMessage('error', '初期化に失敗しました。設定を確認してください。');
        }
    }

    /**
     * UI要素を初期化
     */
    initializeElements() {
        this.elements = {
            // メッセージ
            messageArea: document.getElementById('messageArea'),
            loading: document.getElementById('loading'),
            
            // 会議情報
            currentMeeting: document.getElementById('currentMeeting'),
            refreshMeetingBtn: document.getElementById('refreshMeetingBtn'),
            processVisitorsBtn: document.getElementById('processVisitorsBtn'),
            
            // サービス状態
            sharepointStatus: document.getElementById('sharepointStatus'),
            sharepointStatusText: document.getElementById('sharepointStatusText'),
            powerAutomateStatus: document.getElementById('powerAutomateStatus'),
            powerAutomateStatusText: document.getElementById('powerAutomateStatusText'),
            testConnectionBtn: document.getElementById('testConnectionBtn'),
            
            // 統計
            totalVisitors: document.getElementById('totalVisitors'),
            scheduledVisitors: document.getElementById('scheduledVisitors'),
            completedVisitors: document.getElementById('completedVisitors'),
            refreshStatsBtn: document.getElementById('refreshStatsBtn'),
            
            // 設定
            settingsInfo: document.getElementById('settingsInfo'),
            refreshConfigBtn: document.getElementById('refreshConfigBtn'),
            
            // ログ
            activityLog: document.getElementById('activityLog'),
            clearLogBtn: document.getElementById('clearLogBtn')
        };
    }

    /**
     * イベントリスナーを設定
     */
    setupEventListeners() {
        // ボタンのクリックイベント
        this.elements.refreshMeetingBtn.addEventListener('click', () => this.handleRefreshMeeting());
        this.elements.processVisitorsBtn.addEventListener('click', () => this.handleProcessVisitors());
        this.elements.testConnectionBtn.addEventListener('click', () => this.handleTestConnection());
        this.elements.refreshStatsBtn.addEventListener('click', () => this.handleRefreshStats());
        this.elements.refreshConfigBtn.addEventListener('click', () => this.handleRefreshConfig());
        this.elements.clearLogBtn.addEventListener('click', () => this.handleClearLog());
    }

    /**
     * ログハンドラーを設定
     */
    setupLogHandler() {
        // ログサービスからのメッセージをUIに表示
        if (this.logger && this.logger.addHandler) {
            this.logger.addHandler((level, message, data) => {
                this.addLog(level, message);
            });
        }
    }

    /**
     * サービスを初期化
     */
    async initializeServices() {
        try {
            // 設定サービス
            this.configService = ConfigService.getInstance();
            await this.configService.initialize();
            
            // Outlook サービス
            this.outlookService = new OutlookEventService();
            await this.outlookService.initialize();
            
            // Outlook イベントハンドラーを設定
            this.outlookService.addEventListener('all', (eventType, eventData) => {
                this.handleOutlookEvent(eventType, eventData);
            });
            
            // SharePoint サービス
            this.sharePointService = new SharePointService();
            await this.sharePointService.initialize();
            
            // Power Automate サービス
            this.powerAutomateService = new PowerAutomateService();
            await this.powerAutomateService.initialize();
            
            this.addLog('info', 'すべてのサービスが初期化されました');
            
        } catch (error) {
            this.addLog('error', `サービス初期化エラー: ${error.message}`);
            throw error;
        }
    }

    /**
     * 会議情報を更新
     */
    async refreshMeetingInfo() {
        try {
            if (!this.outlookService) {
                return;
            }

            const meetingData = await this.outlookService.getCurrentEventData();
            this.currentMeetingData = meetingData;
            
            this.displayMeetingInfo(meetingData);
            
            // 外部ユーザーがいる場合、処理ボタンを有効化
            const hasExternalUsers = meetingData && meetingData.externalUsers && meetingData.externalUsers.length > 0;
            this.elements.processVisitorsBtn.disabled = !hasExternalUsers;
            
        } catch (error) {
            this.addLog('error', `会議情報取得エラー: ${error.message}`);
            this.elements.currentMeeting.innerHTML = '<p style="color: #d13438;">会議情報の取得に失敗しました</p>';
        }
    }

    /**
     * 会議情報を表示
     */
    displayMeetingInfo(meetingData) {
        if (!meetingData) {
            this.elements.currentMeeting.innerHTML = '<p style="color: #605e5c;">会議が選択されていません</p>';
            return;
        }

        const startTime = meetingData.startTime ? new Date(meetingData.startTime).toLocaleString('ja-JP') : '未設定';
        const endTime = meetingData.endTime ? new Date(meetingData.endTime).toLocaleString('ja-JP') : '未設定';
        
        let externalUsersHtml = '';
        if (meetingData.externalUsers && meetingData.externalUsers.length > 0) {
            const userTags = meetingData.externalUsers.map(user => 
                `<span class="external-user">${user.emailAddress}</span>`
            ).join('');
            externalUsersHtml = `
                <div class="external-users">
                    <strong>外部ユーザー (${meetingData.externalUsers.length}人):</strong><br>
                    ${userTags}
                </div>
            `;
        } else {
            externalUsersHtml = '<div style="color: #605e5c; font-style: italic;">外部ユーザーはいません</div>';
        }

        this.elements.currentMeeting.innerHTML = `
            <div class="meeting-info">
                <h3>${meetingData.subject || '件名なし'}</h3>
                <div class="meeting-detail">📅 開始: ${startTime}</div>
                <div class="meeting-detail">⏰ 終了: ${endTime}</div>
                <div class="meeting-detail">🆔 会議ID: ${meetingData.meetingId}</div>
                ${externalUsersHtml}
            </div>
        `;
    }

    /**
     * サービス状態を更新
     */
    async refreshServiceStatus() {
        // SharePoint 状態をテスト
        try {
            if (this.sharePointService) {
                const result = await this.sharePointService.testConnection();
                if (result.success) {
                    this.updateServiceStatus('sharepoint', 'connected', `接続済み (${result.itemCount}件)`);
                } else {
                    this.updateServiceStatus('sharepoint', 'disconnected', `エラー: ${result.error}`);
                }
            }
        } catch (error) {
            this.updateServiceStatus('sharepoint', 'disconnected', `接続失敗: ${error.message}`);
        }

        // Power Automate 状態をテスト
        try {
            if (this.powerAutomateService) {
                const result = await this.powerAutomateService.testConnection();
                if (result.success) {
                    this.updateServiceStatus('powerAutomate', 'connected', `接続済み (${result.statusCode})`);
                } else {
                    this.updateServiceStatus('powerAutomate', 'disconnected', `エラー: ${result.error}`);
                }
            }
        } catch (error) {
            this.updateServiceStatus('powerAutomate', 'disconnected', `接続失敗: ${error.message}`);
        }
    }

    /**
     * サービス状態を更新
     */
    updateServiceStatus(service, status, text) {
        const statusElement = this.elements[`${service}Status`];
        const textElement = this.elements[`${service}StatusText`];
        
        if (statusElement && textElement) {
            statusElement.className = `status-indicator status-${status}`;
            textElement.textContent = text;
        }
    }

    /**
     * 設定情報を更新
     */
    async refreshSettings() {
        try {
            if (!this.configService) {
                return;
            }

            const config = this.configService.getConfig();
            const settings = [
                ['SharePoint サイト', config.sharePoint?.siteUrl || '未設定'],
                ['SharePoint リスト', config.sharePoint?.listName || '未設定'],
                ['Power Automate URL', config.powerAutomate?.notificationUrl ? '設定済み' : '未設定'],
                ['内部ドメイン数', config.outlook?.internalDomains?.length || 0],
                ['ログレベル', config.logging?.level || '未設定']
            ];

            this.elements.settingsInfo.innerHTML = settings.map(([label, value]) => `
                <div class="settings-label">${label}:</div>
                <div class="settings-value">${value}</div>
            `).join('');

        } catch (error) {
            this.addLog('error', `設定情報取得エラー: ${error.message}`);
        }
    }

    /**
     * 統計情報を更新
     */
    async refreshStatistics() {
        try {
            if (!this.sharePointService) {
                return;
            }

            const stats = await this.sharePointService.getStatistics(7);
            
            this.elements.totalVisitors.textContent = stats.total;
            this.elements.scheduledVisitors.textContent = stats.scheduled;
            this.elements.completedVisitors.textContent = stats.completed;
            
        } catch (error) {
            this.addLog('error', `統計情報取得エラー: ${error.message}`);
            this.elements.totalVisitors.textContent = '-';
            this.elements.scheduledVisitors.textContent = '-';
            this.elements.completedVisitors.textContent = '-';
        }
    }

    /**
     * Outlook イベントハンドラー
     */
    async handleOutlookEvent(eventType, eventData) {
        try {
            this.addLog('info', `Outlook イベント検知: ${eventType}`);
            
            if (eventData && eventData.externalUsers && eventData.externalUsers.length > 0) {
                // 自動処理を実行
                await this.processVisitors(eventData, eventType);
            }
            
            // UI を更新
            await this.refreshMeetingInfo();
            
        } catch (error) {
            this.addLog('error', `Outlook イベント処理エラー: ${error.message}`);
        }
    }

    /**
     * 来訪者を処理
     */
    async processVisitors(meetingData = null, changeType = 'manual') {
        try {
            if (this.isProcessing) {
                this.showMessage('warning', '処理中です。しばらくお待ちください。');
                return;
            }

            this.isProcessing = true;
            this.showLoading(true);

            const data = meetingData || this.currentMeetingData;
            
            if (!data || !data.externalUsers || data.externalUsers.length === 0) {
                this.showMessage('warning', '処理する外部ユーザーがいません。');
                return;
            }

            this.addLog('info', `来訪者処理を開始: ${data.externalUsers.length}人`);

            // SharePoint にデータを保存/更新
            let sharePointResult;
            if (changeType === 'manual' || changeType === 'created' || changeType === 'recipients_changed') {
                // 新規作成または更新
                sharePointResult = await this.sharePointService.updateVisitorRecords(
                    data.meetingId,
                    data.subject,
                    data.externalUsers,
                    data.startTime,
                    data.endTime
                );
            } else if (changeType === 'deleted') {
                // 削除
                sharePointResult = await this.sharePointService.deleteVisitorRecords(data.meetingId);
            }

            this.addLog('info', `SharePoint 処理完了: ${sharePointResult?.length || 0}件`);

            // Power Automate で通知送信
            if (data.externalUsers.length > 0) {
                const notificationResult = await this.powerAutomateService.sendVisitorNotification(
                    data.meetingId,
                    data.subject,
                    data.externalUsers,
                    data.startTime,
                    data.endTime,
                    changeType === 'manual' ? 'created' : changeType
                );

                this.addLog('info', `通知送信完了: ${notificationResult.notificationsSent}件成功`);

                if (notificationResult.success) {
                    this.showMessage('success', `来訪者処理が完了しました。${data.externalUsers.length}人の外部ユーザーが登録され、通知が送信されました。`);
                } else {
                    this.showMessage('warning', `一部の処理で問題が発生しました。詳細はログを確認してください。`);
                }
            } else {
                this.showMessage('success', '来訪者レコードが更新されました。');
            }

            // 統計を更新
            await this.refreshStatistics();

        } catch (error) {
            this.addLog('error', `来訪者処理エラー: ${error.message}`);
            this.showMessage('error', `処理中にエラーが発生しました: ${error.message}`);
        } finally {
            this.isProcessing = false;
            this.showLoading(false);
        }
    }

    // イベントハンドラー

    async handleRefreshMeeting() {
        this.addLog('info', '会議情報を手動更新');
        await this.refreshMeetingInfo();
    }

    async handleProcessVisitors() {
        await this.processVisitors();
    }

    async handleTestConnection() {
        this.addLog('info', '接続テストを実行');
        this.showLoading(true);
        await this.refreshServiceStatus();
        this.showLoading(false);
    }

    async handleRefreshStats() {
        this.addLog('info', '統計情報を更新');
        this.showLoading(true);
        await this.refreshStatistics();
        this.showLoading(false);
    }

    async handleRefreshConfig() {
        try {
            this.addLog('info', '設定を再読み込み');
            await this.configService.reloadConfig();
            await this.refreshSettings();
            this.showMessage('success', '設定が再読み込みされました。');
        } catch (error) {
            this.addLog('error', `設定再読み込みエラー: ${error.message}`);
            this.showMessage('error', '設定の再読み込みに失敗しました。');
        }
    }

    handleClearLog() {
        this.elements.activityLog.innerHTML = '<div class="log-entry info">ログがクリアされました</div>';
    }

    // ユーティリティメソッド

    /**
     * ローディング表示を制御
     */
    showLoading(show) {
        if (show) {
            this.elements.loading.classList.add('show');
        } else {
            this.elements.loading.classList.remove('show');
        }
    }

    /**
     * メッセージを表示
     */
    showMessage(type, message, duration = 5000) {
        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${type}`;
        messageDiv.textContent = message;
        
        this.elements.messageArea.appendChild(messageDiv);
        
        setTimeout(() => {
            if (messageDiv.parentNode) {
                messageDiv.parentNode.removeChild(messageDiv);
            }
        }, duration);
    }

    /**
     * ログを追加
     */
    addLog(level, message) {
        const timestamp = new Date().toLocaleTimeString('ja-JP');
        const logEntry = document.createElement('div');
        logEntry.className = `log-entry ${level}`;
        logEntry.textContent = `[${timestamp}] ${message}`;
        
        this.elements.activityLog.appendChild(logEntry);
        
        // スクロールを最下部に
        this.elements.activityLog.scrollTop = this.elements.activityLog.scrollHeight;
        
        // ログエントリ数を制限
        const entries = this.elements.activityLog.children;
        if (entries.length > 100) {
            this.elements.activityLog.removeChild(entries[0]);
        }
    }
}

// アプリケーションを開始
const app = new TaskpaneApp();

// DOM読み込み完了後に初期化
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        app.initialize().catch(console.error);
    });
} else {
    app.initialize().catch(console.error);
}

// グローバルアクセス用（デバッグなど）
window.lobbyApp = app;