class ConfigManager {
    constructor() {
        this.store = null;
        this.initPromise = null;
    }

    async init() {
        if (this.initPromise) {
            return this.initPromise;
        }

        this.initPromise = this._doInit();
        return this.initPromise;
    }

    async _doInit() {
        try {
            // const { default: Store } = await import('electron-store');
            const importStore = new Function('return import("electron-store")');
            const { default: Store } = await importStore();

            this.store = new Store({
                defaults: {
                    windowBounds: { width: 1200, height: 800 },
                    theme: 'light',
                    language: 'zh-CN'
                }
            });
            console.log('ConfigManager 初始化完成');
        } catch (error) {
            console.error('ConfigManager 初始化失败:', error);
            throw error;
        }
    }

    async get(key) {
        await this.init();
        return this.store.get(key);
    }

    async set(key, value) {
        await this.init();
        this.store.set(key, value);
    }

    async getAll() {
        await this.init();
        return this.store.store;
    }
    async getPath(){
        await this.init();
        return this.store.path;
    }
}

// 导出单例
export default new ConfigManager();
