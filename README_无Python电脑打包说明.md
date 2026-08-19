# 给没有 Python 的 Windows 电脑使用：EXE 打包说明

领导电脑没有 Python 时，不要发脚本版。应先打包成：

```text
PPTDiffTool_v0_9.exe
```

再发给领导测试。

## 推荐方式：用 GitHub Actions 云端打包

你本人不需要 Windows 电脑，也不需要领导安装 Python。

### 第 1 步：新建 GitHub 仓库

在 GitHub 新建一个空仓库，例如：

```text
ppt-diff-tool-build
```

### 第 2 步：上传本文件夹全部内容

把本文件夹中的所有文件上传到仓库根目录，至少包括：

```text
ppt_diff_tool.py
PPTDiffTool_v0_9.spec
requirements.txt
README_Windows.md
Windows测试清单.md
.github/workflows/build-windows-exe.yml
```

### 第 3 步：运行 Actions

进入 GitHub 仓库：

```text
Actions
→ Build Windows EXE
→ Run workflow
```

等待 2–5 分钟。

### 第 4 步：下载构建产物

构建完成后，在 workflow 页面底部下载 artifact：

```text
PPTDiffTool_v0_9_Windows_EXE
```

解压后得到：

```text
PPTDiffTool_v0_9.exe
README_Windows.md
Windows测试清单.md
```

把这三个文件发给领导即可。

## 领导使用方式

领导双击：

```text
PPTDiffTool_v0_9.exe
```

不需要安装 Python。

## 注意事项

1. Windows 可能提示“未知发布者”，这是因为 exe 没有代码签名。
2. 如被公司杀毒软件拦截，需要 IT 放行，或改用 Python 脚本版。
3. 当前 exe 是本地工具，不上传 PPT，不联网。
4. 仍然只支持 `.pptx`，不支持 `.ppt`。
