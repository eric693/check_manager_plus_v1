// config.js

const API_CONFIG = {
  // 正式環境的 API URL
  apiUrl: "https://script.google.com/macros/s/AKfycbzdHwEUXTL00fq38AUzAOA1pdTa2g-IzgWZtdgjC8zYblJZs7-nU8tVDDTfQNVAtYVI/exec",
  
  // 新增回呼網址
  redirectUrl: "https://eric693.github.io/check_manager_plus_v1/"
  // 你也可以在這裡加入其他設定，例如：
  // timeout: 5000,
  // version: 'v4.3.4'
};
// 👇 新增：為了兼容性，同時定義全域變數 apiUrl
const apiUrl = API_CONFIG.apiUrl;
