// Enhanced utility functions
//HANDLES THE 429 ERRORS
interface RetryableError {
    headers?: {
        get?: (header: string) => string | null;
    };
    error?: {
        innerError?: {
            retryAfter?: string | number;
        };
    };
    status?: number;
}

export function delay(ms: number): Promise<void> & { cancel: () => void } {
    let timeoutId: number;
    const promise = new Promise<void>((resolve, reject) => {
        timeoutId = setTimeout(resolve, ms);
    });
    (promise as any).cancel = () => clearTimeout(timeoutId);
    return promise as Promise<void> & { cancel: () => void };
}

export function getRetryAfterMs(error: unknown, fallbackMs = 3000): number {
    const retryableError = error as RetryableError;
    
    const retryAfter = 
        retryableError?.headers?.get?.('Retry-After') ??
        retryableError?.error?.innerError?.retryAfter ??
        null;

    if (retryAfter !== null) {
        const seconds = Number(retryAfter);
        if (!Number.isNaN(seconds) && seconds > 0) {
            return Math.min(seconds * 1000, 30000); //at 30 seconds
        }
    }
    
    // Exponential backoff for HTTP errors
    if (retryableError?.status && retryableError.status >= 500) {
        return fallbackMs * 2;
    }
    
    return fallbackMs;
}

export async function graphRequestWithRetry<T>(
    request: () => Promise<T>,
    options: {
        maxRetries?: number;
        baseDelay?: number;
        maxDelay?: number;
        shouldRetry?: (error: unknown) => boolean;
        onRetry?: (error: unknown, attempt: number) => void;
    } = {}
): Promise<T> {
    const {
        maxRetries = 3,
        baseDelay = 1000,
        maxDelay = 10000,
        shouldRetry = (error: unknown) => {
            const err = error as RetryableError;
            return !err.status || err.status >= 500 || err.status === 429;
        },
        onRetry
    } = options;

    let lastError: unknown;
    
    for (let attempt = 0; attempt <= maxRetries; attempt++) {
        try {
            return await request();
        } catch (error) {
            lastError = error;
            
            if (attempt === maxRetries || !shouldRetry(error)) {
                throw error;
            }
            
            const baseRetryAfter = getRetryAfterMs(error, baseDelay);
            // Add jitter to prevent thundering herd
            const jitter = Math.random() * 0.3 * baseRetryAfter;
            const delayMs = Math.min(baseRetryAfter + jitter, maxDelay);
            
            onRetry?.(error, attempt);
            await delay(delayMs);
        }
    }
    
    throw lastError;
}

export interface ConcurrencyOptions<T, R> {
    limit: number;
    worker: (item: T, index: number) => Promise<R>;
    onProgress?: (completed: number, total: number) => void;
    onError?: (error: unknown, item: T, index: number) => void;
}

export class Semaphore {
    private permits: number;
    private waitQueue: (() => void)[] = [];
    
    constructor(permits: number) {
        this.permits = permits;
    }
    
    async acquire(): Promise<void> {
        if (this.permits > 0) {
            this.permits--;
            return;
        }
        
        return new Promise<void>((resolve) => {
            this.waitQueue.push(resolve);
        });
    }
    
    release(): void {
        this.permits++;
        if (this.waitQueue.length > 0) {
            const resolve = this.waitQueue.shift()!;
            this.permits--;
            resolve();
        }
    }
}

export async function runWithConcurrencyLimit<T, R>(
    items: T[],
    options: ConcurrencyOptions<T, R>
): Promise<R[]> {
    const { limit, worker, onProgress, onError } = options;
    const results: (R | Error)[] = new Array(items.length);
    let completed = 0;
    
    const semaphore = new Semaphore(limit);
    
    const promises = items.map(async (item, index) => {
        await semaphore.acquire();
        
        try {
            const result = await worker(item, index);
            results[index] = result;
            return result;
        } catch (error) {
            results[index] = error as Error;
            onError?.(error, item, index);
            throw error;
        } finally {
            completed++;
            onProgress?.(completed, items.length);
            semaphore.release();
        }
    });
    
    await Promise.allSettled(promises);
    
    // Filter out errors and throw if any failed
    const errors = results.filter((r): r is Error => r instanceof Error);
    if (errors.length > 0) {
        throw new AggregateError(errors, `${errors.length} of ${items.length} tasks failed`);
    }
    
    return results as R[];
}
