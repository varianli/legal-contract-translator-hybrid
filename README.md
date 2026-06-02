# Legal Contract Translator Hybrid

一个面向法律合同的 Word 文档中译英工具，目标是在调用大模型进行上下文翻译的同时，尽量保留 `.docx` 原有格式，包括段落结构、表格、粗体、下划线、高亮、页眉页脚和字体设置。

## Features

- 支持 OpenAI 兼容接口和 DeepSeek API。
- 使用 Markdown 结构保留条款层级和上下文，再结合 Word run 结构映射格式。
- 支持多文件翻译，每个文件独立显示翻译进度和复核进度。
- 支持 Windows 和 macOS 启动脚本。
- 导出英文 Word 文档，并将中间 Markdown、checklist、JSON 明细放入过程文件夹。
- 最终自检查会扫描中文残留和机器标记残留，例如 `<!-- META:... -->`。
- API Key 只保存在本地配置文件中，相关配置文件已加入 `.gitignore`。

## Quick Start

### Windows

进入 `复合方法` 文件夹，双击：

```text
一键启动.bat
```

### macOS

进入 `复合方法` 文件夹，双击：

```text
一键启动.command
```

如果 macOS 提示无法打开，可在终端中进入该目录后执行：

```bash
chmod +x 一键启动.command
./一键启动.command
```

## Basic Usage

1. 选择 API 类型。
2. 输入 API Key、Base URL 和模型名称。
3. 添加一个或多个 `.docx` 文件。
4. 选择输出目录。
5. 点击开始翻译。

DeepSeek 默认配置：

```text
Base URL: https://api.deepseek.com
Model: deepseek-v4-flash
```

## Project Structure

```text
复合方法/
  hybrid_markdown_run_translator.py      # Windows / shared translation core
  hybrid_markdown_run_translator_mac.py  # macOS dedicated GUI
  一键启动.bat                            # Windows launcher
  一键启动.command                        # macOS launcher
  requirements.txt                       # Windows dependencies
  requirements_mac.txt                   # macOS dependencies
```

## Notes

This is an AI-assisted legal translation tool. Outputs should be reviewed by a qualified legal professional before formal submission, signing, or filing.
