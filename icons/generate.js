/**
 * 生成插件图标 (PNG via Canvas)
 * 在插件安装时运行，生成各尺寸图标
 * 注：本文件仅用于说明，实际图标已内置
 */

// 如需重新生成图标，可在浏览器控制台运行以下代码：
function generateIcon(size) {
  const canvas = document.createElement('canvas');
  canvas.width = size;
  canvas.height = size;
  const ctx = canvas.getContext('2d');
  
  // 背景渐变
  const grad = ctx.createLinearGradient(0, 0, size, size);
  grad.addColorStop(0, '#FF6B00');
  grad.addColorStop(1, '#FF9500');
  ctx.fillStyle = grad;
  ctx.roundRect(0, 0, size, size, size * 0.15);
  ctx.fill();
  
  // 表格图标
  ctx.strokeStyle = 'white';
  ctx.lineWidth = size * 0.06;
  ctx.fillStyle = 'white';
  
  const p = size * 0.2;
  const w = size * 0.6;
  const h = size * 0.55;
  const rowH = h / 3;
  
  // 表格框
  ctx.strokeRect(p, p + rowH * 0.3, w, h);
  // 横线
  ctx.beginPath();
  ctx.moveTo(p, p + rowH * 0.3 + rowH);
  ctx.lineTo(p + w, p + rowH * 0.3 + rowH);
  ctx.moveTo(p, p + rowH * 0.3 + rowH * 2);
  ctx.lineTo(p + w, p + rowH * 0.3 + rowH * 2);
  // 竖线
  ctx.moveTo(p + w * 0.4, p + rowH * 0.3);
  ctx.lineTo(p + w * 0.4, p + rowH * 0.3 + h);
  ctx.stroke();
  
  return canvas.toDataURL('image/png');
}
