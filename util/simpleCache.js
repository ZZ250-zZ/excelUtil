class SimpleCache {
    constructor() {
        this.cache = new Map();
        this.startAutoCleanup()
    }

    set(key, value) {
        const currentTime = new Date();
        this.cache.set(key, { value, timestamp: currentTime });
    }

    get(key) {
        const entry = this.cache.get(key);
        if (!entry) return undefined;

        const currentTime = new Date();
        const TWO_HOURS = 2 * 60 * 60 * 1000;
        if (currentTime - entry.timestamp > TWO_HOURS) {
            this.cache.delete(key);
            return undefined;
        }
        return entry.value;
    }

    startAutoCleanup(intervalInMilliseconds = 30 * 60 * 1000) {
        if (this.checkInterval) clearInterval(this.checkInterval);
        this.checkInterval = setInterval(() => this.cleanup(), intervalInMilliseconds);
    }

    stopAutoCleanup() {
        if (this.checkInterval) {
            clearInterval(this.checkInterval);
            this.checkInterval = null;
        }
    }

    cleanup() {
        const currentTime = new Date();
        for (const [key, entry] of this.cache.entries()) {
            if (currentTime - entry.timestamp > 60 * 60 * 1000) {
                this.cache.delete(key);
            }
        }
    }
}

module.exports = SimpleCache; // 导出SimpleCache类，以便在其他文件中使用