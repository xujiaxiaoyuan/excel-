// Excel中的操作需要在Office.onReady回调函数中执行
Office.onReady(function(info) {
  // 如果Office.js加载成功，则向选定单元格中写入文本
  if (info.host === Office.HostType.Excel) {
    $(document).ready(function() {
      $('#mark-color-button').click(markColor);
    });
  }
});

function markColor() {
  // 获取当前选定的单元格
  Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.load('values');
    return context.sync()
        .then(function() {
          // 将大于100且小于500的数值标记为红色，大于500的数值标记为蓝色
          for (var i = 0; i < range.values.length; i++) {
            for (var j = 0; j < range.values[i].length; j++) {
              if (range.values[i][j] > 100 && range.values[i][j] < 500) {
                range.getCell(i, j).format.font.color = 'red';
              }
              if (range.values[i][j] >= 500) {
                range.getCell(i, j).format.font.color = 'blue';
              }
            }
          }
          return context.sync();
        });
  });
}
