// config.js

const API_CONFIG = {
  // 正式環境的 API URL
  apiUrl: "https://script.google.com/macros/s/AKfycbzdHwEUXTL00fq38AUzAOA1pdTa2g-IzgWZtdgjC8zYblJZs7-nU8tVDDTfQNVAtYVI/exec",
  
  // 新增回呼網址
  redirectUrl: "https://eric693.github.io/check_manager_plus_v1/",

  // LINE Login Channel ID（公開資訊，會出現在登入網址中）。
  // 前端直接組登入網址，省掉一次呼叫後端 getLoginUrl 的等待
  lineChannelId: "2008779593"
  // 你也可以在這裡加入其他設定，例如：
  // timeout: 5000,
  // version: 'v4.5.2'
};
// 👇 新增：為了兼容性，同時定義全域變數 apiUrl
const apiUrl = API_CONFIG.apiUrl;
