# XlsxSimplify
## 一款基于 XlsxJs 的在线预览excel插件

### 案例
https://hulalalalala.github.io/XlsxSimplify/dist/index.html

### 用法
```
var nContainer = document.getElementById('container') // 嵌入表格的dom节点
var xlsx = new Plus.Xlsx(nContainer) // 实例化插件类，传入DOM节点
xlsx.readFile('./test.xlsx') // 渲染文件 可传 文件路径 或者 文件流FileData对象
```
