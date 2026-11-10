// 在表A的 AirScript / JS脚本编辑器中创建“文档共享脚本”，粘贴本文件。
// 只读，不修改表A。适用于普通工作表API，不适用于多维表的记录API。
const a = Context.argv || {};
function columnNumber(s) {
  if (!/^[A-Z]{1,3}$/.test(s)) throw new Error("列名必须为A、B、AA等");
  let n = 0;
  for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}
const sheet = Application.Worksheets.Item(a.sheet_name);
if (a.mode === 'filter') {
  const filterCol = columnNumber(a.filter_column), firstRow = Number(a.source_start_row);
  if (filterCol > 16384 || !Number.isInteger(firstRow) || firstRow < 1 || firstRow > 1048576 || typeof a.filter_value !== 'string' || !a.filter_value.length) throw new Error('筛选配置错误');
  const used = sheet.UsedRange;
  const endRow = used.Row + used.Rows.Count - 1;
  const width = used.Column + used.Columns.Count - 1;
  if (filterCol > width) throw new Error('筛选列超出表A实际使用范围');
  const scanCount = Math.max(0, endRow - firstRow + 1);
  if (scanCount > Number(a.max_cells)) throw new Error('筛选列扫描超过max_cells，请调整读取上限或清理无用格式');
  function letters(n) { let s=''; while(n){const r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26)}return s; }
  function matrix(ref, height, columns) {
    let value=sheet.Range(ref).Value2;
    if (!Array.isArray(value)) value=[[value]];
    if (value.length!==height || !value.every(row=>Array.isArray(row)&&row.length===columns)) throw new Error('读取区域不完整');
    return value;
  }
  const rows=[];
  const matches=scanCount ? matrix(a.filter_column+firstRow+':'+a.filter_column+endRow,scanCount,1) : [];
  for(let i=0;i<matches.length;i++){
    const value=matches[i][0];
    if(String(value===null||value===undefined?'':value)!==a.filter_value) continue;
    if(rows.length+1>Number(a.max_rows) || scanCount+(rows.length+1)*width>Number(a.max_cells)) throw new Error('筛选结果超过读取或导出上限，未截断输出');
    const rowNo=firstRow+i;
    const data=matrix('A'+rowNo+':'+letters(width)+rowNo,1,width)[0];
    if(String(data[filterCol-1]===null||data[filterCol-1]===undefined?'':data[filterCol-1])!==a.filter_value) throw new Error('读取期间筛选值发生变化，请重试');
    rows.push([rowNo,JSON.stringify(data)]);
  }
  return JSON.stringify({version:1,complete:true,headers:['源行号','行数据'],rows,row_count:rows.length,
    column_count:width, selection:{filter_column:a.filter_column,filter_value:a.filter_value,source_start_row:firstRow},
    range:'A'+firstRow+':'+letters(width)+endRow,captured_at:new Date().toISOString()});
}
if (a.mode === 'cells') {
  if (!Array.isArray(a.cells) || !a.cells.length || a.cells.length > 500) throw new Error('请选择1至500组单元格');
  if (a.cells.length > Number(a.max_cells)) throw new Error('所选单元格超过最大读取单元格数');
  const seen = new Set(), rows = [];
  for (const cell of a.cells) {
    if (!Number.isInteger(cell.row) || cell.row < 1 || cell.row > 1048576 ||
        !Number.isInteger(cell.column) || cell.column < 1 || cell.column > 16384) throw new Error('行列号无效');
    let n = cell.column, col = '';
    while (n) { const r = (n - 1) % 26; col = String.fromCharCode(65 + r) + col; n = Math.floor((n - 1) / 26); }
    const ref = col + cell.row;
    if (seen.has(ref)) continue;
    seen.add(ref);
    let value = sheet.Range(ref).Value2;
    if (Array.isArray(value)) value = value[0][0];
    rows.push([ref, value === undefined ? null : value]);
  }
  return JSON.stringify({version:1, complete:true, headers:['单元格','值'], rows,
    row_count:rows.length, range:'指定单元格', captured_at:new Date().toISOString()});
}
const first = columnNumber(a.first_column);
const last = columnNumber(a.last_column);
const header = Number(a.header_row);
if (!Number.isInteger(header) || header < 1 || last < first || last > 16384) {
  throw new Error("读取范围配置错误");
}
const used = sheet.UsedRange;
const endRow = Math.max(header, used.Row + used.Rows.Count - 1);
const width = last - first + 1;
if ((endRow - header + 1) * width > Number(a.max_cells)) {
  throw new Error("范围超过max_cells，请缩小列范围或清理尾部无用格式后重试");
}
const address = a.first_column + header + ":" + a.last_column + endRow;
let values = sheet.Range(address).Value2;
if (!Array.isArray(values)) values = [[values]];
if (!values.every(row => Array.isArray(row) && row.length === width)) {
  throw new Error("接口没有返回完整的二维区域");
}
if (values.length !== endRow - header + 1) throw new Error("返回行数不完整");
const headers = values[0];
const rows = values.slice(1).filter(row => row.some(v => v !== null && v !== ""));
// 单次区域读取减少多次分页时行移动产生的漏行，仍不等同于源端事务快照。
return JSON.stringify({version: 1, complete: true, headers, rows,
  row_count: rows.length, range: address, captured_at: new Date().toISOString()});
