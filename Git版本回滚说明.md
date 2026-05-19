# Git 版本回滚说明

这个文件夹已经初始化为 Git 仓库。当前提交是你回退后的稳定基线。

## 日常保存一个版本

在这个文件夹空白处右键打开终端，执行：

```powershell
git status
git add .
git commit -m "说明这次改了什么"
```

## 查看历史版本

```powershell
git log --oneline
```

## 回滚单个文件

先用 `git log --oneline` 找到想回到的版本号，然后执行：

```powershell
git restore --source 版本号 -- 新方案/new_scheme_translator.py
git restore --source 版本号 -- 复合方法/hybrid_markdown_run_translator.py
```

## 回滚整个文件夹

这个会把当前未提交修改全部丢弃，请确认不需要当前改动后再执行：

```powershell
git reset --hard 版本号
```

## 当前忽略内容

`.gitignore` 已排除：

- API key 配置文件
- `.venv` 虚拟环境
- `__pycache__` 缓存
- `输出` 文件夹里的翻译结果
