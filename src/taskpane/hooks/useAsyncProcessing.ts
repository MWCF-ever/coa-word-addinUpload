import { useState, useCallback, useRef } from 'react';
import axios from 'axios';
import { API_BASE_URL } from '../../types';

interface TaskProgress {
    current: number;
    total: number;
    status: string;
}

interface ProcessingResult {
    batch_data?: any[];
    processed_files?: string[];
    failed_files?: any[];
    from_cache?: boolean;
}

export const useAsyncProcessing = () => {
    const [taskProgress, setTaskProgress] = useState<TaskProgress | null>(null);
    const [isProcessing, setIsProcessing] = useState(false);
    const intervalRef = useRef<ReturnType<typeof setInterval> | null>(null);

    const processAsync = useCallback(async (params: {
        compound_id: string;
        template_id: string;
        force_reprocess?: boolean;
    }): Promise<string> => {
        try {
            const response = await axios.post(
                `${API_BASE_URL}/api/documents/process-directory-async`,
                params
            );

            if (response.data.success && response.data.data.task_id) {
                setIsProcessing(true);
                return response.data.data.task_id;
            } else {
                throw new Error('Failed to start async processing');
            }
        } catch (error) {
            console.error('Error starting async processing:', error);
            throw error;
        }
    }, []);

    const checkTaskStatus = useCallback(async (taskId: string): Promise<ProcessingResult | null> => {
        return new Promise((resolve, reject) => {
            let pollCount = 0;
            const maxPolls = 60; // 最多轮询60次（5分钟）

            const pollStatus = async () => {
                try {
                    pollCount++;
                    
                    const response = await axios.get(
                        `${API_BASE_URL}/api/documents/task-status/${taskId}`
                    );

                    if (!response.data.success) {
                        throw new Error('Failed to get task status');
                    }

                    const data = response.data.data;

                    if (data.state === 'PROGRESS') {
                        setTaskProgress({
                            current: data.current,
                            total: data.total,
                            status: data.status
                        });
                    } else if (data.state === 'SUCCESS') {
                        if (intervalRef.current) {
                            clearInterval(intervalRef.current);
                            intervalRef.current = null;
                        }
                        setIsProcessing(false);
                        setTaskProgress(null);
                        resolve(data.result);
                    } else if (data.state === 'FAILURE') {
                        if (intervalRef.current) {
                            clearInterval(intervalRef.current);
                            intervalRef.current = null;
                        }
                        setIsProcessing(false);
                        setTaskProgress(null);
                        reject(new Error(data.error || 'Task failed'));
                    }

                    // 超时检查
                    if (pollCount >= maxPolls) {
                        if (intervalRef.current) {
                            clearInterval(intervalRef.current);
                            intervalRef.current = null;
                        }
                        setIsProcessing(false);
                        setTaskProgress(null);
                        reject(new Error('Task timeout'));
                    }
                } catch (error) {
                    if (intervalRef.current) {
                        clearInterval(intervalRef.current);
                        intervalRef.current = null;
                    }
                    setIsProcessing(false);
                    setTaskProgress(null);
                    reject(error);
                }
            };

            // 立即执行第一次检查
            pollStatus();

            // 设置轮询间隔
            intervalRef.current = setInterval(pollStatus, 5000); // 每5秒检查一次
        });
    }, []);

    const cancelTask = useCallback(async (taskId: string) => {
        try {
            await axios.delete(`${API_BASE_URL}/api/documents/cancel-task/${taskId}`);
            
            if (intervalRef.current) {
                clearInterval(intervalRef.current);
                intervalRef.current = null;
            }
            
            setIsProcessing(false);
            setTaskProgress(null);
        } catch (error) {
            console.error('Error cancelling task:', error);
        }
    }, []);

    // 组件卸载时清理
    const cleanup = useCallback(() => {
        if (intervalRef.current) {
            clearInterval(intervalRef.current);
            intervalRef.current = null;
        }
    }, []);

    return {
        processAsync,
        checkTaskStatus,
        cancelTask,
        cleanup,
        isProcessing,
        taskProgress
    };
};