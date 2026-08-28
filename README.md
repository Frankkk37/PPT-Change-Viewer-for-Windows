# PPT Diff MVP v0.9 Windows 版

这是基于 v0.9 的 Windows 测试版，用于让 Windows 电脑直接测试。

## 定位

当前版本是 MVP，不做版本树、不做多人协作、不做 PowerPoint 插件。

核心流程：

```text
选择旧版 PPT
选择新版 PPT
点击比较
生成 HTML / Markdown / JSON 报告
```

## 文件说明

```text
ppt_diff_tool.py          主程序
启动PPTDiff.bat           Windows 双击启动
README_Windows.md         本说明
Windows测试清单.md         给测试者用的验收清单
```

## 使用方式

### 方法一：双击启动

双击：

```text
启动PPTDiff.bat
```

然后在界面中选择：

```text
旧版 PPT（File A）
新版 PPT（File B）
输出文件夹
```

点击：

```text
比较两个 PPT
```

完成后会自动打开 HTML 报告，也可以点击：

```text
打开 HTML 报告
打开 Markdown 报告
打开输出文件夹
```

### 方法二：命令行

```bat
python ppt_diff_tool.py "旧版.pptx" "新版.pptx" -o diff_output
```

如果需要检测格式/布局变化：

```bat
python ppt_diff_tool.py "旧版.pptx" "新版.pptx" -o diff_output --detect-format
```

## Python 要求

如果直接运行脚本版，需要 Windows 电脑已安装 Python 3。

如果领导电脑没有 Python，请不要使用脚本版；请按 `README_无Python电脑打包说明.md` 先打包成 `.exe`。

如果双击无法启动，通常是没有安装 Python，或安装时没有勾选：

```text
Add python.exe to PATH
```

Python 下载地址：

```text
https://www.python.org/downloads/windows/
```

## 输出结果

每次比较会生成：

```text
xxx.diff.html
xxx.diff.md
xxx.diff.json
```

其中：

```text
HTML：主要给人看
Markdown：方便复制到微信/企业微信/飞书
JSON：机器可读，后续版本树可用
```

## 当前支持

- .pptx 文件
- 新增页
- 删除页
- 移动页
- 修改页
- 短句级文字变化
- 图片变化
- 形状数量变化
- 表格/图表容器数量变化
- 页码变化过滤
- HTML / Markdown / JSON 报告

## 当前不支持

- .ppt 老格式
- PowerPoint 插件
- 自动合并
- 版本树
- 云协作
- 精确表格单元格 diff
- 视觉截图 diff

## 测试建议

优先测试：

```text
1. 只改两个 typo
2. 插入一页
3. 删除第一页
4. 移动第一页到第二页
5. 修改一句较长文字，检查短句级 diff 是否好读
```
