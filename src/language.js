const english = {
  preview: 'Preview document', close: 'Close preview', loading: 'Loading document…',
  error: 'Preview unavailable. Check the file, access and CORS settings.',
  page: 'Page', sheet: 'Worksheet',
  readOnly: 'Read only', cellAddress: 'Cell address', cellValue: 'Cell value or cached formula', go: 'Go to cell', invalidAddress: 'Enter a cell inside this worksheet.',
  zoomIn: 'Zoom in', zoomOut: 'Zoom out', fitWidth: 'Fit width', expand: 'Expand workspace', collapse: 'Return to page', hiddenCell: 'This cell is hidden in the document.'
};
const simplified = {
  preview: '预览文档', close: '关闭预览', loading: '正在加载文档…',
  error: '无法预览。请检查文件、访问权限和跨域设置。',
  page: '页码', sheet: '工作表', readOnly: '只读预览', cellAddress: '单元格地址', cellValue: '单元格值或已保存的公式', go: '定位', invalidAddress: '请输入此工作表内的单元格地址。',
  zoomIn: '放大', zoomOut: '缩小', fitWidth: '适合宽度', expand: '展开工作区', collapse: '返回页面', hiddenCell: '此单元格在原文档中已隐藏。'
};
const traditional = {
  preview: '預覽文件', close: '關閉預覽', loading: '正在載入文件…',
  error: '無法預覽。請檢查檔案、存取權限及跨來源設定。',
  page: '頁碼', sheet: '工作表', readOnly: '唯讀預覽', cellAddress: '儲存格位置', cellValue: '儲存格值或已儲存的公式', go: '定位', invalidAddress: '請輸入此工作表內的儲存格位置。',
  zoomIn: '放大', zoomOut: '縮小', fitWidth: '符合寬度', expand: '展開工作區', collapse: '返回頁面', hiddenCell: '此儲存格在原文件中已隱藏。'
};

export function documentLabels(locale = 'en') {
  return { ...(/^zh-(?:tw|hk|hant)/i.test(locale) ? traditional : /^zh/i.test(locale) ? simplified : english) };
}
