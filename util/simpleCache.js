const MAX_CACHE_SIZE = 100; // 设置缓存条数 上限

// 半小时
const HALF_HOURS = 60 * 1000 * 30;
// 1h
const ONE_HOURS = 60 * 1000 * 60;
// 2h
const TWO_HOURS = ONE_HOURS * 2;

// 缓存的过期时间
const TIMEOUT = ONE_HOURS
// 定时任务的时间间隔
const TIME_INTERVAL = HALF_HOURS

class SimpleCache {
    constructor() {
        this.cache = new Map();
        this.startAutoCleanup()
    }

    set(key, value) {
        const currentTime = new Date();

        // 如果缓存达到大小上限，删除最早的条目
        if (this.cache.size >= MAX_CACHE_SIZE) {
            throw new Error("内存泄露啦，请联系管理员.");
        }
        this.cache.set(key, { value, timestamp: currentTime });
    }

    get(key) {
        const entry = this.cache.get(key);
        if (!entry) return undefined;

        const currentTime = new Date();
        if (currentTime - entry.timestamp > TIMEOUT) {
            this.cache.delete(key);
            return undefined;
        }
        return entry.value;
    }

    delete(key) {
        this.cache.delete(key);
    }

    // 半小时检查过期状态
    startAutoCleanup(intervalInMilliseconds = TIME_INTERVAL) {
        if (this.checkInterval) clearInterval(this.checkInterval);
        this.checkInterval = setInterval(() => this.cleanup(), intervalInMilliseconds);
    }

    stopAutoCleanup() {
        if (this.checkInterval) {
            clearInterval(this.checkInterval);
            this.checkInterval = null;
        }
    }

    // 删除超过 1h 的key
    cleanup() {
        const currentTime = new Date();
        for (const [key, entry] of this.cache.entries()) {
            if (currentTime - entry.timestamp > TIMEOUT) {
                this.cache.delete(key);
            }
        }
    }

    // 安全的清空缓存
    clear() {
        this.stopAutoCleanup();
        this.cache.clear();
    }
}

module.exports = SimpleCache; // 导出SimpleCache类，以便在其他文件中使用