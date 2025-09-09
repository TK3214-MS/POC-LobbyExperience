/**
 * SharePointService
 * SharePoint Onlineリストとの連携を担当するサービス
 */

import { ConfigService } from './ConfigService.js';
import { LoggingService } from './LoggingService.js';

export class SharePointService {
    constructor() {
        this.config = ConfigService.getInstance();
        this.logger = LoggingService.getInstance();
        this.accessToken = null;
        this.tokenExpiry = null;
        this.isInitialized = false;
        this.retryConfig = {
            maxRetries: 3,
            retryDelay: 1000,
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

            const sharePointConfig = this.config.getConfig().sharePoint;
            if (!sharePointConfig) {
                throw new Error('SharePoint configuration is missing');
            }

            if (!sharePointConfig.siteUrl || !sharePointConfig.listName) {
                throw new Error('SharePoint siteUrl and listName are required');
            }

            // アクセストークンを取得
            await this.getAccessToken();
            
            this.isInitialized = true;
            this.logger.info('SharePointService initialized successfully');
        } catch (error) {
            this.logger.error('Failed to initialize SharePointService', error);
            throw error;
        }
    }

    /**
     * アクセストークンを取得
     */
    async getAccessToken() {
        try {
            // トークンが有効かチェック
            if (this.accessToken && this.tokenExpiry && new Date() < this.tokenExpiry) {
                return this.accessToken;
            }

            const config = this.config.getConfig().sharePoint;
            
            // Office.jsのSSOを使用してトークンを取得
            if (typeof Office !== 'undefined' && Office.auth && Office.auth.getAccessToken) {
                const token = await new Promise((resolve, reject) => {
                    Office.auth.getAccessToken({
                        resource: 'https://graph.microsoft.com'
                    }, (result) => {
                        if (result.status === 'succeeded') {
                            resolve(result.value);
                        } else {
                            reject(new Error(result.error.message));
                        }
                    });
                });

                this.accessToken = token;
                // トークンの有効期限を設定（通常1時間、安全のため50分に設定）
                this.tokenExpiry = new Date(Date.now() + 50 * 60 * 1000);
                
                return this.accessToken;
            }

            // フォールバック: MSALを使用
            if (typeof PublicClientApplication !== 'undefined') {
                return await this.getMsalToken();
            }

            throw new Error('No authentication method available');

        } catch (error) {
            this.logger.error('Failed to get access token', error);
            throw error;
        }
    }

    /**
     * MSALを使用してトークンを取得
     */
    async getMsalToken() {
        try {
            const config = this.config.getConfig().sharePoint;
            
            const msalConfig = {
                auth: {
                    clientId: config.clientId,
                    authority: `https://login.microsoftonline.com/${config.tenantId}`
                }
            };

            const msalInstance = new PublicClientApplication(msalConfig);
            
            const request = {
                scopes: ['https://graph.microsoft.com/.default'],
                account: msalInstance.getAllAccounts()[0]
            };

            const response = await msalInstance.acquireTokenSilent(request);
            
            this.accessToken = response.accessToken;
            this.tokenExpiry = new Date(response.expiresOn);
            
            return this.accessToken;

        } catch (error) {
            this.logger.error('MSAL token acquisition failed', error);
            throw error;
        }
    }

    /**
     * SharePointリストに来訪者情報を作成
     */
    async createVisitorRecord(meetingId, meetingTitle, externalUsers, startTime, endTime) {
        try {
            const records = [];
            
            for (const user of externalUsers) {
                const record = {
                    MeetingId: meetingId,
                    MeetingTitle: meetingTitle,
                    VisitorEmail: user.emailAddress,
                    VisitorName: user.name || user.emailAddress.split('@')[0],
                    StartTime: this.formatDateTime(startTime),
                    EndTime: this.formatDateTime(endTime),
                    Status: 'Scheduled',
                    CreatedDate: this.formatDateTime(new Date()),
                    ModifiedDate: this.formatDateTime(new Date())
                };
                
                const createdRecord = await this.createListItem(record);
                records.push(createdRecord);
            }
            
            this.logger.info(`Created ${records.length} visitor records for meeting ${meetingId}`);
            return records;

        } catch (error) {
            this.logger.error('Failed to create visitor records', error);
            throw error;
        }
    }

    /**
     * SharePointリストのアイテムを作成
     */
    async createListItem(itemData) {
        return await this.executeWithRetry(async () => {
            const config = this.config.getConfig().sharePoint;
            const token = await this.getAccessToken();
            
            const url = `${config.siteUrl}/_api/web/lists/getByTitle('${config.listName}')/items`;
            
            const response = await fetch(url, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose',
                    'X-RequestDigest': await this.getRequestDigest()
                },
                body: JSON.stringify(itemData)
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const data = await response.json();
            return data.d;
        });
    }

    /**
     * 会議IDで来訪者レコードを取得
     */
    async getVisitorRecordsByMeetingId(meetingId) {
        return await this.executeWithRetry(async () => {
            const config = this.config.getConfig().sharePoint;
            const token = await this.getAccessToken();
            
            const filter = `MeetingId eq '${meetingId}'`;
            const url = `${config.siteUrl}/_api/web/lists/getByTitle('${config.listName}')/items?$filter=${encodeURIComponent(filter)}`;
            
            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Accept': 'application/json;odata=verbose'
                }
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const data = await response.json();
            return data.d.results;
        });
    }

    /**
     * 来訪者レコードを更新
     */
    async updateVisitorRecords(meetingId, meetingTitle, externalUsers, startTime, endTime) {
        try {
            // 既存のレコードを取得
            const existingRecords = await this.getVisitorRecordsByMeetingId(meetingId);
            
            // 既存のレコードを削除
            for (const record of existingRecords) {
                await this.deleteListItem(record.Id);
            }
            
            // 新しいレコードを作成
            if (externalUsers.length > 0) {
                const newRecords = await this.createVisitorRecord(
                    meetingId, 
                    meetingTitle, 
                    externalUsers, 
                    startTime, 
                    endTime
                );
                
                this.logger.info(`Updated visitor records for meeting ${meetingId}`);
                return newRecords;
            }
            
            this.logger.info(`Removed all visitor records for meeting ${meetingId} (no external users)`);
            return [];

        } catch (error) {
            this.logger.error('Failed to update visitor records', error);
            throw error;
        }
    }

    /**
     * 来訪者レコードを削除
     */
    async deleteVisitorRecords(meetingId) {
        try {
            const existingRecords = await this.getVisitorRecordsByMeetingId(meetingId);
            
            for (const record of existingRecords) {
                await this.deleteListItem(record.Id);
            }
            
            this.logger.info(`Deleted ${existingRecords.length} visitor records for meeting ${meetingId}`);
            return existingRecords.length;

        } catch (error) {
            this.logger.error('Failed to delete visitor records', error);
            throw error;
        }
    }

    /**
     * SharePointリストのアイテムを削除
     */
    async deleteListItem(itemId) {
        return await this.executeWithRetry(async () => {
            const config = this.config.getConfig().sharePoint;
            const token = await this.getAccessToken();
            
            const url = `${config.siteUrl}/_api/web/lists/getByTitle('${config.listName}')/items(${itemId})`;
            
            const response = await fetch(url, {
                method: 'DELETE',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Accept': 'application/json;odata=verbose',
                    'X-RequestDigest': await this.getRequestDigest(),
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'DELETE'
                }
            });

            if (!response.ok && response.status !== 404) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            return response.ok;
        });
    }

    /**
     * RequestDigestを取得
     */
    async getRequestDigest() {
        try {
            const config = this.config.getConfig().sharePoint;
            const token = await this.getAccessToken();
            
            const url = `${config.siteUrl}/_api/contextinfo`;
            
            const response = await fetch(url, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Accept': 'application/json;odata=verbose'
                }
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const data = await response.json();
            return data.d.GetContextWebInformation.FormDigestValue;

        } catch (error) {
            this.logger.error('Failed to get request digest', error);
            // フォールバック値を返す
            return 'dummy-digest';
        }
    }

    /**
     * 日時をSharePoint形式にフォーマット
     */
    formatDateTime(date) {
        if (!date) return null;
        
        if (typeof date === 'string') {
            date = new Date(date);
        }
        
        return date.toISOString();
    }

    /**
     * リストの構造を検証
     */
    async validateListStructure() {
        try {
            const config = this.config.getConfig().sharePoint;
            const token = await this.getAccessToken();
            
            const url = `${config.siteUrl}/_api/web/lists/getByTitle('${config.listName}')/fields`;
            
            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Accept': 'application/json;odata=verbose'
                }
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const data = await response.json();
            const fields = data.d.results;
            
            const requiredFields = [
                'MeetingId',
                'MeetingTitle', 
                'VisitorEmail',
                'VisitorName',
                'StartTime',
                'EndTime',
                'Status'
            ];
            
            const missingFields = requiredFields.filter(field => 
                !fields.some(f => f.InternalName === field)
            );
            
            if (missingFields.length > 0) {
                this.logger.warn('Missing required fields in SharePoint list', missingFields);
                return { valid: false, missingFields };
            }
            
            this.logger.info('SharePoint list structure validation passed');
            return { valid: true, missingFields: [] };

        } catch (error) {
            this.logger.error('Failed to validate list structure', error);
            return { valid: false, error: error.message };
        }
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
                
                this.logger.warn(`Operation failed, retrying in ${delay}ms (attempt ${retryCount + 1}/${this.retryConfig.maxRetries})`, error);
                
                await new Promise(resolve => setTimeout(resolve, delay));
                return await this.executeWithRetry(operation, retryCount + 1);
            }
            
            this.logger.error(`Operation failed after ${this.retryConfig.maxRetries} retries`, error);
            throw error;
        }
    }

    /**
     * 接続テスト
     */
    async testConnection() {
        try {
            const config = this.config.getConfig().sharePoint;
            const token = await this.getAccessToken();
            
            const url = `${config.siteUrl}/_api/web/lists/getByTitle('${config.listName}')`;
            
            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Accept': 'application/json;odata=verbose'
                }
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const data = await response.json();
            this.logger.info('SharePoint connection test successful');
            
            return {
                success: true,
                listTitle: data.d.Title,
                itemCount: data.d.ItemCount
            };

        } catch (error) {
            this.logger.error('SharePoint connection test failed', error);
            return {
                success: false,
                error: error.message
            };
        }
    }

    /**
     * 統計情報を取得
     */
    async getStatistics(dateRange = 7) {
        try {
            const config = this.config.getConfig().sharePoint;
            const token = await this.getAccessToken();
            
            const cutoffDate = new Date();
            cutoffDate.setDate(cutoffDate.getDate() - dateRange);
            
            const filter = `CreatedDate ge datetime'${cutoffDate.toISOString()}'`;
            const url = `${config.siteUrl}/_api/web/lists/getByTitle('${config.listName}')/items?$filter=${encodeURIComponent(filter)}&$select=Status,CreatedDate`;
            
            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Accept': 'application/json;odata=verbose'
                }
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const data = await response.json();
            const records = data.d.results;
            
            const stats = {
                total: records.length,
                scheduled: records.filter(r => r.Status === 'Scheduled').length,
                completed: records.filter(r => r.Status === 'Completed').length,
                cancelled: records.filter(r => r.Status === 'Cancelled').length,
                dateRange: dateRange
            };
            
            this.logger.info('Retrieved SharePoint statistics', stats);
            return stats;

        } catch (error) {
            this.logger.error('Failed to get statistics', error);
            throw error;
        }
    }

    /**
     * サービスを破棄
     */
    dispose() {
        try {
            this.accessToken = null;
            this.tokenExpiry = null;
            this.isInitialized = false;
            this.logger.info('SharePointService disposed');
        } catch (error) {
            this.logger.error('Error disposing SharePointService', error);
        }
    }
}