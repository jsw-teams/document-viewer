const english = {
  preview: 'Preview document', close: 'Close preview', download: 'Download original', loading: 'Loading document…',
  error: 'Preview unavailable. Download the original, or check the file and CORS settings.',
  previous: 'Previous page', next: 'Next page', page: 'Page', sheet: 'Worksheet',
  limited: 'Preview is limited to 1,000 rows and 100 columns per worksheet.',
  pptLimited: 'Basic legacy PPT preview: animations, complex charts, master inheritance and some drawing effects are not supported.'
};
const simplified = {
  preview: '预览文档', close: '关闭预览', download: '下载原文件', loading: '正在加载文档…',
  error: '无法预览。请下载原文件，或检查文件和跨域设置。',
  previous: '上一页', next: '下一页', page: '页码', sheet: '工作表', limited: '每个工作表最多预览 1,000 行、100 列。',
  pptLimited: '旧版 PPT 为基础预览：暂不支持动画、复杂图表、母版继承及部分绘图效果。'
};
const traditional = {
  preview: '預覽文件', close: '關閉預覽', download: '下載原檔案', loading: '正在載入文件…',
  error: '無法預覽。請下載原檔案，或檢查檔案及跨來源設定。',
  previous: '上一頁', next: '下一頁', page: '頁碼', sheet: '工作表', limited: '每個工作表最多預覽 1,000 列、100 欄。',
  pptLimited: '舊版 PPT 為基礎預覽：暫不支援動畫、複雜圖表、母片繼承及部分繪圖效果。'
};

export function documentLabels(locale = 'en') {
  return { ...(/^zh-(?:tw|hk|hant)/i.test(locale) ? traditional : /^zh/i.test(locale) ? simplified : english) };
}
