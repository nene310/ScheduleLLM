/**
 * Runtime Configuration
 * 
 * INSTRUCTIONS:
 * 1. Copy this file to 'config.js' (which is git-ignored)
 * 2. Replace the placeholder values below with your actual API keys
 * 3. OR use the 'inject_env.js' script to generate config.js from environment variables
 */

window.AppConfig = {
    // API Key for Alibaba Cloud Qwen (Tongyi Qianwen)
    // Get it from: https://bailian.console.aliyun.com/
    apiKey: "YOUR_API_KEY_HERE",

    // API Base URL (Default: Aliyun Compatible Mode)
    baseUrl: "https://dashscope.aliyuncs.com/compatible-mode/v1",

    // Default Model
    model: "qwen-turbo"
};
