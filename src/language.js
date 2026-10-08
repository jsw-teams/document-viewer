const english = {
  preview: 'Preview document', close: 'Close preview', loading: 'Loading document…',
  error: 'Preview unavailable. Check the file, access and CORS settings.',
  page: 'Page', sheet: 'Worksheet',
  limited: 'Preview is limited to 1,000 rows and 100 columns per worksheet.'
};
const simplified = {
  preview: '预览文档', close: '关闭预览', loading: '正在加载文档…',
  error: '无法预览。请检查文件、访问权限和跨域设置。',
  page: '页码', sheet: '工作表', limited: '每个工作表最多预览 1,000 行、100 列。'
};
const traditional = {
  preview: '預覽文件', close: '關閉預覽', loading: '正在載入文件…',
  error: '無法預覽。請檢查檔案、存取權限及跨來源設定。',
  page: '頁碼', sheet: '工作表', limited: '每個工作表最多預覽 1,000 列、100 欄。'
};

export function documentLabels(locale = 'en') {
  return { ...(/^zh-(?:tw|hk|hant)/i.test(locale) ? traditional : /^zh/i.test(locale) ? simplified : english) };
}
