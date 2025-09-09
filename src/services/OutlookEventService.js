/**
 * OutlookEventService
 * Outlook予定表のイベント監視と外部ユーザー抽出を担当するサービス
 */

import { ConfigService } from './ConfigService.js';
import { LoggingService } from './LoggingService.js';

export class OutlookEventService {
    constructor() {
        this.config = ConfigService.getInstance();
        this.logger = LoggingService.getInstance();
        this.eventHandlers = new Map();
        this.isInitialized = false;
    }

    /**
     * サービスを初期化
     */
    async initialize() {
        try {
            if (this.isInitialized) {
                return;
            }

            // Office.jsが利用可能かチェック
            if (typeof Office === 'undefined') {
                throw new Error('Office.js is not available');
            }

            await Office.onReady();
            
            // Outlook固有の機能をチェック
            if (!Office.context.mailbox) {
                throw new Error('Mailbox context is not available');
            }

            // イベントハンドラーを設定
            this.setupEventHandlers();
            
            this.isInitialized = true;
            this.logger.info('OutlookEventService initialized successfully');
        } catch (error) {
            this.logger.error('Failed to initialize OutlookEventService', error);
            throw error;
        }
    }

    /**
     * イベントハンドラーを設定
     */
    setupEventHandlers() {
        try {
            // アイテム変更イベントの監視
            if (Office.context.mailbox.item) {
                // 作成モードまたは読み取りモードでのイベント監視
                this.monitorCurrentItem();
            }

            // メールボックスレベルでのイベント監視
            this.monitorMailboxEvents();
            
            this.logger.info('Event handlers setup completed');
        } catch (error) {
            this.logger.error('Failed to setup event handlers', error);
        }
    }

    /**
     * 現在のアイテム（会議）を監視
     */
    monitorCurrentItem() {
        const item = Office.context.mailbox.item;
        
        if (item && item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            // 受信者変更の監視
            if (item.requiredAttendees) {
                item.requiredAttendees.addHandlerAsync(
                    Office.EventType.RecipientsChanged,
                    this.handleRecipientsChanged.bind(this)
                );
            }

            if (item.optionalAttendees) {
                item.optionalAttendees.addHandlerAsync(
                    Office.EventType.RecipientsChanged,
                    this.handleRecipientsChanged.bind(this)
                );
            }

            // 件名変更の監視
            if (item.subject) {
                item.subject.addHandlerAsync(
                    Office.EventType.ItemChanged,
                    this.handleSubjectChanged.bind(this)
                );
            }

            // 開始・終了時刻変更の監視
            if (item.start) {
                item.start.addHandlerAsync(
                    Office.EventType.ItemChanged,
                    this.handleTimeChanged.bind(this)
                );
            }

            if (item.end) {
                item.end.addHandlerAsync(
                    Office.EventType.ItemChanged,
                    this.handleTimeChanged.bind(this)
                );
            }
        }
    }

    /**
     * メールボックスレベルでのイベント監視
     */
    monitorMailboxEvents() {
        // アポイントメントの作成・更新・削除を監視
        if (Office.context.mailbox.addHandlerAsync) {
            Office.context.mailbox.addHandlerAsync(
                Office.EventType.AppointmentTimeChanged,
                this.handleAppointmentChanged.bind(this)
            );
        }
    }

    /**
     * 受信者変更ハンドラー
     */
    async handleRecipientsChanged(eventArgs) {
        try {
            this.logger.info('Recipients changed event triggered');
            
            const item = Office.context.mailbox.item;
            const eventData = await this.extractEventData(item);
            
            if (eventData && eventData.externalUsers.length > 0) {
                // 外部ユーザーが存在する場合、処理を実行
                await this.processEventChange('recipients_changed', eventData);
            }
        } catch (error) {
            this.logger.error('Error handling recipients changed', error);
        }
    }

    /**
     * 件名変更ハンドラー
     */
    async handleSubjectChanged(eventArgs) {
        try {
            this.logger.info('Subject changed event triggered');
            
            const item = Office.context.mailbox.item;
            const eventData = await this.extractEventData(item);
            
            if (eventData && eventData.externalUsers.length > 0) {
                await this.processEventChange('subject_changed', eventData);
            }
        } catch (error) {
            this.logger.error('Error handling subject changed', error);
        }
    }

    /**
     * 時刻変更ハンドラー
     */
    async handleTimeChanged(eventArgs) {
        try {
            this.logger.info('Time changed event triggered');
            
            const item = Office.context.mailbox.item;
            const eventData = await this.extractEventData(item);
            
            if (eventData && eventData.externalUsers.length > 0) {
                await this.processEventChange('time_changed', eventData);
            }
        } catch (error) {
            this.logger.error('Error handling time changed', error);
        }
    }

    /**
     * アポイントメント変更ハンドラー
     */
    async handleAppointmentChanged(eventArgs) {
        try {
            this.logger.info('Appointment changed event triggered');
            
            const item = Office.context.mailbox.item;
            const eventData = await this.extractEventData(item);
            
            if (eventData && eventData.externalUsers.length > 0) {
                await this.processEventChange('appointment_changed', eventData);
            }
        } catch (error) {
            this.logger.error('Error handling appointment changed', error);
        }
    }

    /**
     * 現在の会議から外部ユーザーとイベントデータを抽出
     */
    async extractEventData(item) {
        return new Promise((resolve, reject) => {
            try {
                if (!item || item.itemType !== Office.MailboxEnums.ItemType.Appointment) {
                    resolve(null);
                    return;
                }

                const eventData = {
                    meetingId: item.itemId || this.generateMeetingId(),
                    subject: '',
                    startTime: null,
                    endTime: null,
                    externalUsers: [],
                    allAttendees: []
                };

                let pendingOperations = 0;
                let completedOperations = 0;

                const checkCompletion = () => {
                    if (completedOperations === pendingOperations) {
                        // 外部ユーザーの抽出
                        eventData.externalUsers = this.filterExternalUsers(eventData.allAttendees);
                        resolve(eventData);
                    }
                };

                // 件名取得
                if (item.subject) {
                    pendingOperations++;
                    item.subject.getAsync((result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            eventData.subject = result.value;
                        }
                        completedOperations++;
                        checkCompletion();
                    });
                }

                // 開始時刻取得
                if (item.start) {
                    pendingOperations++;
                    item.start.getAsync((result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            eventData.startTime = result.value;
                        }
                        completedOperations++;
                        checkCompletion();
                    });
                }

                // 終了時刻取得
                if (item.end) {
                    pendingOperations++;
                    item.end.getAsync((result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            eventData.endTime = result.value;
                        }
                        completedOperations++;
                        checkCompletion();
                    });
                }

                // 必須出席者取得
                if (item.requiredAttendees) {
                    pendingOperations++;
                    item.requiredAttendees.getAsync((result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            eventData.allAttendees.push(...result.value);
                        }
                        completedOperations++;
                        checkCompletion();
                    });
                }

                // 任意出席者取得
                if (item.optionalAttendees) {
                    pendingOperations++;
                    item.optionalAttendees.getAsync((result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            eventData.allAttendees.push(...result.value);
                        }
                        completedOperations++;
                        checkCompletion();
                    });
                }

                // 非同期操作がない場合
                if (pendingOperations === 0) {
                    resolve(eventData);
                }

            } catch (error) {
                this.logger.error('Error extracting event data', error);
                reject(error);
            }
        });
    }

    /**
     * 外部ユーザーをフィルタリング
     */
    filterExternalUsers(attendees) {
        try {
            const config = this.config.getConfig();
            const internalDomains = config.outlook?.internalDomains || [];
            const excludeRooms = config.outlook?.excludeRooms || true;

            return attendees.filter(attendee => {
                if (!attendee.emailAddress) {
                    return false;
                }

                const email = attendee.emailAddress.toLowerCase();
                
                // 会議室を除外
                if (excludeRooms && this.isRoomEmail(email)) {
                    return false;
                }

                // 内部ドメインをチェック
                const isInternal = internalDomains.some(domain => 
                    email.endsWith(`@${domain.toLowerCase()}`)
                );

                return !isInternal;
            });
        } catch (error) {
            this.logger.error('Error filtering external users', error);
            return [];
        }
    }

    /**
     * 会議室のメールアドレスかどうかを判定
     */
    isRoomEmail(email) {
        // 一般的な会議室メールアドレスのパターン
        const roomPatterns = [
            /^room\d+/i,
            /^conference/i,
            /^meeting/i,
            /^boardroom/i,
            /^conf-/i,
            /-room$/i,
            /^resource-/i
        ];

        return roomPatterns.some(pattern => pattern.test(email));
    }

    /**
     * 会議IDを生成
     */
    generateMeetingId() {
        return `meeting_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    }

    /**
     * イベント変更を処理
     */
    async processEventChange(changeType, eventData) {
        try {
            this.logger.info(`Processing event change: ${changeType}`, eventData);

            // イベントハンドラーを呼び出し
            if (this.eventHandlers.has(changeType)) {
                const handlers = this.eventHandlers.get(changeType);
                for (const handler of handlers) {
                    await handler(eventData);
                }
            }

            // 全てのハンドラーを呼び出し
            if (this.eventHandlers.has('all')) {
                const handlers = this.eventHandlers.get('all');
                for (const handler of handlers) {
                    await handler(changeType, eventData);
                }
            }

        } catch (error) {
            this.logger.error(`Error processing event change: ${changeType}`, error);
            throw error;
        }
    }

    /**
     * イベントハンドラーを登録
     */
    addEventListener(eventType, handler) {
        if (!this.eventHandlers.has(eventType)) {
            this.eventHandlers.set(eventType, []);
        }
        this.eventHandlers.get(eventType).push(handler);
    }

    /**
     * イベントハンドラーを削除
     */
    removeEventListener(eventType, handler) {
        if (this.eventHandlers.has(eventType)) {
            const handlers = this.eventHandlers.get(eventType);
            const index = handlers.indexOf(handler);
            if (index > -1) {
                handlers.splice(index, 1);
            }
        }
    }

    /**
     * 手動でイベントデータを取得
     */
    async getCurrentEventData() {
        try {
            const item = Office.context.mailbox.item;
            return await this.extractEventData(item);
        } catch (error) {
            this.logger.error('Error getting current event data', error);
            throw error;
        }
    }

    /**
     * サービスを破棄
     */
    dispose() {
        try {
            // イベントハンドラーをクリア
            this.eventHandlers.clear();
            this.isInitialized = false;
            this.logger.info('OutlookEventService disposed');
        } catch (error) {
            this.logger.error('Error disposing OutlookEventService', error);
        }
    }
}