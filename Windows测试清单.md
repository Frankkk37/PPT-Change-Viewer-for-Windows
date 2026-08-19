# Windows 测试清单

## 测试环境

请记录：

```text
Windows 版本：
Python 版本：
PowerPoint 版本：
测试 PPT 文件页数：
```

## 测试 1：启动

操作：

```text
双击 启动PPTDiff.bat
```

预期：

```text
能打开 GUI 界面
```

如果不能打开，请记录终端报错截图。

## 测试 2：只改 typo

操作：

```text
旧版和新版页数相同，只改一页中的两个错别字
```

预期：

```text
Added: 0
Removed: 0
Moved: 0
Modified: 1
HTML / Markdown 均能显示具体文字变化
```

## 测试 3：插入一页

操作：

```text
在中间插入一页
```

预期：

```text
Added: 1
Moved: 后续页面整体后移
Modified: 不应大面积误报
```

## 测试 4：删除第一页

操作：

```text
删除 P1
```

预期：

```text
Removed: 1
Moved: 后续页面整体前移
Modified: 不应大面积误报
```

## 测试 5：移动页面

操作：

```text
将 P1 移到 P2 后
```

预期：

```text
Moved: 2
Modified: 0
```

## 测试 6：报告按钮

操作：

```text
比较完成后点击：
打开 HTML 报告
打开 Markdown 报告
打开输出文件夹
```

预期：

```text
三个按钮均能正常打开
```

## 测试 7：中文路径

操作：

```text
将工具和 PPT 放在包含中文的路径中，例如：
桌面/测试文件夹/希格项目/
```

预期：

```text
能正常选择文件、生成报告、打开报告
```

## 测试反馈

请记录：

```text
是否能启动：
是否能完成比较：
HTML 是否好读：
Markdown 是否能打开：
有无明显误报：
错误截图：
```
