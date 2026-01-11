const fs = require('fs');
const path = require('path');

/**
 * Configuration Generator
 * 
 * This script reads environment variables and generates the 'config.js' file.
 * It is designed for:
 * 1. CI/CD pipelines where secrets are injected via env vars (e.g. process.env.API_KEY)
 * 2. Docker deployments using --env-file
 * 3. Local development using a .env file
 */

// 1. Load .env file if it exists (Simple parsing)
const envPath = path.join(__dirname, '.env');
if (fs.existsSync(envPath)) {
    console.log('Loading .env file...');
    const envContent = fs.readFileSync(envPath, 'utf-8');
    envContent.split('\n').forEach(line => {
        const match = line.match(/^\s*([\w_]+)\s*=\s*(.*)?\s*$/);
        if (match) {
            const key = match[1];
            let value = match[2] || '';
            // Remove quotes if present
            if (value.startsWith('"') && value.endsWith('"')) value = value.slice(1, -1);
            if (value.startsWith("'") && value.endsWith("'")) value = value.slice(1, -1);
            
            if (!process.env[key]) {
                process.env[key] = value;
            }
        }
    });
}

// 2. Extract specific variables
const config = {
    llmApiUrl: process.env.LLM_API_URL || "/api/llm",
    model: process.env.LLM_MODEL || "qwen-flash"
};

// 3. Generate config.js content
const fileContent = `/**
 * Auto-generated configuration file
 * Generated at: ${new Date().toISOString()}
 */

window.AppConfig = {
    llmApiUrl: "${config.llmApiUrl}",
    model: "${config.model}"
};
`;

// 4. Write to config.js
const outputPath = path.join(__dirname, 'config.js');
fs.writeFileSync(outputPath, fileContent);

console.log('Successfully generated config.js');
console.log('LLM API URL configured:', config.llmApiUrl);
