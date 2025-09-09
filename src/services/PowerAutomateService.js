/**
 * PowerAutomateService
 * Power Automate APIを使用した外部ユーザー通知サービス
 */

import { ConfigService } from './ConfigService.js';
import { LoggingService } from './LoggingService.js';

export class PowerAutomateService {
    constructor() {
        this.config = ConfigService.getInstance();
        this.logger = LoggingService.getInstance();
        this.isInitialized = false;
        this.retryConfig = {
            maxRetries: 3,
            retryDelay: 2000,
            backoffMultiplier: 2
        };
    }

    /**
     * サービスを初期化
     */
    async initialize() {
        try {
            if (this.isInitialized) {
                return;
            }

            const powerAutomateConfig = this.config.getConfig().powerAutomate;
            if (!powerAutomateConfig) {
                throw new Error('Power Automate configuration is missing');
            }

            if (!powerAutomateConfig.notificationUrl) {
                throw new Error('Power Automate notification URL is required');
            }

            // URLの有効性をチェック
            try {
                new URL(powerAutomateConfig.notificationUrl);
            } catch (error) {
                throw new Error('Invalid Power Automate notification URL');
            }

            this.isInitialized = true;
            this.logger.info('PowerAutomateService initialized successfully');
        } catch (error) {
            this.logger.error('Failed to initialize PowerAutomateService', error);
            throw error;
        }
    }

    /**
     * 外部ユーザーに来訪通知を送信
     */
    async sendVisitorNotification(meetingId, meetingTitle, externalUsers, startTime, endTime, notificationType = 'created') {
        try {
            if (!this.isInitialized) {
                await this.initialize();
            }

            if (!externalUsers || externalUsers.length === 0) {
                this.logger.info('No external users to notify');
                return { success: true, notificationsSent: 0 };
            }

            const notifications = [];
            
            // 各外部ユーザーに対して通知を送信
            for (const user of externalUsers) {
                try {
                    const notification = await this.sendSingleNotification({
                        meetingId,
                        meetingTitle,
                        visitorEmail: user.emailAddress,
                        visitorName: user.name || user.emailAddress.split('@')[0],
                        startTime,
                        endTime,
                        notificationType
                    });
                    
                    notifications.push({
                        email: user.emailAddress,
                        success: true,
                        notificationId: notification.notificationId
                    });
                    
                } catch (error) {
                    this.logger.error(`Failed to send notification to ${user.emailAddress}`, error);
                    notifications.push({
                        email: user.emailAddress,
                        success: false,
                        error: error.message
                    });
                }
            }

            const successCount = notifications.filter(n => n.success).length;
            const failureCount = notifications.filter(n => !n.success).length;

            this.logger.info(`Visitor notifications sent: ${successCount} success, ${failureCount} failures`);

            return {
                success: failureCount === 0,
                notificationsSent: successCount,
                failures: failureCount,
                details: notifications
            };

        } catch (error) {
            this.logger.error('Failed to send visitor notifications', error);
            throw error;
        }
    }

    /**
     * 単一のユーザーに通知を送信
     */
    async sendSingleNotification(notificationData) {
        return await this.executeWithRetry(async () => {
            const config = this.config.getConfig().powerAutomate;
            const timeout = config.timeout || 30000;

            const payload = {
                meetingId: notificationData.meetingId,
                meetingTitle: notificationData.meetingTitle,
                visitorEmail: notificationData.visitorEmail,
                visitorName: notificationData.visitorName,
                startTime: this.formatDateTime(notificationData.startTime),
                endTime: this.formatDateTime(notificationData.endTime),
                notificationType: notificationData.notificationType,
                timestamp: new Date().toISOString(),
                source: 'LobbyExperienceAddin'
            };

            this.logger.info(`Sending notification to Power Automate for ${notificationData.visitorEmail}`, {
                meetingId: notificationData.meetingId,
                type: notificationData.notificationType
            });

            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), timeout);

            try {
                const response = await fetch(config.notificationUrl, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'Accept': 'application/json'
                    },
                    body: JSON.stringify(payload),
                    signal: controller.signal
                });

                clearTimeout(timeoutId);

                if (!response.ok) {
                    const errorText = await response.text().catch(() => 'Unknown error');
                    throw new Error(`HTTP ${response.status}: ${errorText}`);
                }

                // レスポンスを解析（Power Automateは様々な形式で応答する可能性がある）
                let responseData;
                try {
                    const contentType = response.headers.get('content-type');
                    if (contentType && contentType.includes('application/json')) {
                        responseData = await response.json();
                    } else {
                        responseData = await response.text();
                    }
                } catch (parseError) {
                    // レスポンスが解析できない場合でも成功とする
                    responseData = 'Success';
                }

                this.logger.info(`Notification sent successfully to ${notificationData.visitorEmail}`);

                return {
                    success: true,
                    notificationId: this.generateNotificationId(),
                    response: responseData,
                    timestamp: new Date().toISOString()
                };

            } catch (error) {
                clearTimeout(timeoutId);
                
                if (error.name === 'AbortError') {
                    throw new Error(`Request timeout after ${timeout}ms`);
                }
                
                throw error;
            }
        });
    }

    /**
     * 会議更新通知を送信
     */
    async sendMeetingUpdateNotification(meetingId, meetingTitle, externalUsers, startTime, endTime, changeType) {
        try {
            const notificationType = this.mapChangeTypeToNotification(changeType);
            
            return await this.sendVisitorNotification(
                meetingId,
                meetingTitle,
                externalUsers,
                startTime,
                endTime,
                notificationType
            );
        } catch (error) {
            this.logger.error('Failed to send meeting update notification', error);
            throw error;
        }
    }

    /**
     * 会議キャンセル通知を送信
     */
    async sendMeetingCancellationNotification(meetingId, meetingTitle, externalUsers) {
        try {
            return await this.sendVisitorNotification(
                meetingId,
                meetingTitle,
                externalUsers,
                null,
                null,
                'cancelled'
            );
        } catch (error) {
            this.logger.error('Failed to send meeting cancellation notification', error);
            throw error;
        }
    }

    /**
     * バルク通知送信（複数の会議に対して）
     */
    async sendBulkNotifications(meetings) {
        try {
            if (!meetings || meetings.length === 0) {
                return { success: true, processed: 0 };
            }

            const results = [];
            
            for (const meeting of meetings) {
                try {
                    const result = await this.sendVisitorNotification(
                        meeting.meetingId,
                        meeting.meetingTitle,
                        meeting.externalUsers,
                        meeting.startTime,
                        meeting.endTime,
                        meeting.notificationType || 'created'
                    );
                    
                    results.push({
                        meetingId: meeting.meetingId,
                        success: true,
                        result
                    });
                } catch (error) {
                    this.logger.error(`Failed to process bulk notification for meeting ${meeting.meetingId}`, error);
                    results.push({
                        meetingId: meeting.meetingId,
                        success: false,
                        error: error.message
                    });
                }
            }

            const successCount = results.filter(r => r.success).length;
            const failureCount = results.filter(r => !r.success).length;

            this.logger.info(`Bulk notifications processed: ${successCount} success, ${failureCount} failures`);

            return {
                success: failureCount === 0,
                processed: results.length,
                successes: successCount,
                failures: failureCount,
                details: results
            };

        } catch (error) {
            this.logger.error('Failed to send bulk notifications', error);
            throw error;
        }
    }

    /**
     * Power Automate接続テスト
     */
    async testConnection() {
        try {
            if (!this.isInitialized) {
                await this.initialize();
            }

            const testPayload = {
                test: true,
                timestamp: new Date().toISOString(),
                source: 'LobbyExperienceAddin',
                message: 'Connection test from Outlook Add-in'
            };

            const config = this.config.getConfig().powerAutomate;
            const timeout = config.timeout || 30000;

            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), timeout);

            try {
                const response = await fetch(config.notificationUrl, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'Accept': 'application/json'
                    },
                    body: JSON.stringify(testPayload),
                    signal: controller.signal
                });

                clearTimeout(timeoutId);

                const success = response.ok;
                const statusCode = response.status;
                
                let responseData;
                try {
                    const contentType = response.headers.get('content-type');
                    if (contentType && contentType.includes('application/json')) {
                        responseData = await response.json();
                    } else {
                        responseData = await response.text();
                    }
                } catch {
                    responseData = 'No response data';
                }

                this.logger.info(`Power Automate connection test completed: ${success ? 'SUCCESS' : 'FAILED'}`, {
                    status: statusCode,
                    response: responseData
                });

                return {
                    success,
                    statusCode,
                    response: responseData,
                    timestamp: new Date().toISOString()
                };

            } catch (error) {
                clearTimeout(timeoutId);
                
                if (error.name === 'AbortError') {
                    throw new Error(`Connection test timeout after ${timeout}ms`);
                }
                
                throw error;
            }

        } catch (error) {
            this.logger.error('Power Automate connection test failed', error);
            return {
                success: false,
                error: error.message,
                timestamp: new Date().toISOString()
            };
        }
    }

    /**
     * 変更タイプを通知タイプにマッピング
     */
    mapChangeTypeToNotification(changeType) {
        const mapping = {
            'recipients_changed': 'updated',
            'subject_changed': 'updated',
            'time_changed': 'updated',
            'appointment_changed': 'updated',
            'created': 'created',
            'deleted': 'cancelled'
        };
        
        return mapping[changeType] || 'updated';
    }

    /**
     * 日時をISO文字列にフォーマット
     */
    formatDateTime(date) {
        if (!date) return null;
        
        if (typeof date === 'string') {
            return new Date(date).toISOString();
        }
        
        return date.toISOString();
    }

    /**
     * 通知IDを生成
     */
    generateNotificationId() {
        return `notification_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    }

    /**
     * リトライ機能付きで関数を実行
     */
    async executeWithRetry(operation, retryCount = 0) {
        try {
            return await operation();
        } catch (error) {
            if (retryCount < this.retryConfig.maxRetries) {
                const delay = this.retryConfig.retryDelay * 
                    Math.pow(this.retryConfig.backoffMultiplier, retryCount);
                
                this.logger.warn(`Power Automate operation failed, retrying in ${delay}ms (attempt ${retryCount + 1}/${this.retryConfig.maxRetries})`, error);
                
                await new Promise(resolve => setTimeout(resolve, delay));
                return await this.executeWithRetry(operation, retryCount + 1);
            }
            
            this.logger.error(`Power Automate operation failed after ${this.retryConfig.maxRetries} retries`, error);
            throw error;
        }
    }

    /**
     * 通知履歴を取得（ローカルストレージから）
     */
    getNotificationHistory(limit = 50) {
        try {
            const historyKey = 'powerAutomate_notification_history';
            const historyStr = localStorage.getItem(historyKey);
            
            if (!historyStr) {
                return [];
            }
            
            const history = JSON.parse(historyStr);
            return Array.isArray(history) ? history.slice(0, limit) : [];
            
        } catch (error) {
            this.logger.error('Failed to get notification history', error);
            return [];
        }
    }

    /**
     * 通知履歴を保存
     */
    saveNotificationHistory(notification) {
        try {
            const historyKey = 'powerAutomate_notification_history';
            const maxHistorySize = 100;
            
            let history = this.getNotificationHistory(maxHistorySize);
            
            // 新しい通知を履歴の先頭に追加
            history.unshift({
                ...notification,
                timestamp: new Date().toISOString()
            });
            
            // 履歴サイズを制限
            if (history.length > maxHistorySize) {
                history = history.slice(0, maxHistorySize);
            }
            
            localStorage.setItem(historyKey, JSON.stringify(history));
            
        } catch (error) {
            this.logger.error('Failed to save notification history', error);
        }
    }

    /**
     * 通知履歴をクリア
     */
    clearNotificationHistory() {
        try {
            const historyKey = 'powerAutomate_notification_history';
            localStorage.removeItem(historyKey);
            this.logger.info('Notification history cleared');
        } catch (error) {
            this.logger.error('Failed to clear notification history', error);
        }
    }

    /**
     * サービスを破棄
     */
    dispose() {
        try {
            this.isInitialized = false;
            this.logger.info('PowerAutomateService disposed');
        } catch (error) {
            this.logger.error('Error disposing PowerAutomateService', error);
        }
    }
}