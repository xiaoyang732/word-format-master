# 开发验证与发布

Python 3.10+；先安装 `requirements.txt`。必要测试和验证工具在 Git 中管理，运行包与开发仓库使用独立收集规则。

```powershell
python -m unittest discover -s dev/tests
python dev/scripts/validate_project.py
python dev/scripts/smoke_test.py
```

构建输出必须是不存在的新 ZIP 路径：

```powershell
python dev/tools/build_release.py --output dev/artifacts/word-format-master-1.2.4.zip
python dev/tools/build_release.py --output dev/artifacts/word-format-master-1.2.4-json.zip --template-mode json
python dev/tools/validate_release.py dev/artifacts/word-format-master-1.2.4-json.zip
```

构建入口收集运行代码、规范、配置及模板，校验登记的模板摘要与来源/使用说明。JSON 模式将 DOCX 编码成既有加载器支持的 `.docx.json`。构建后先在临时目录校验全部文件摘要，再清空源码导入影响和模板缓存，使用解压包自己的 CLI 运行默认报告列表与准备流程；通过后才写出最终 ZIP。

包内 `release-manifest.json` 记录版本、提交、工作区是否有未提交修改和所有运行文件摘要。运行包不包含开发测试、历史文档和个人输出。构建阶段不向外部服务上传文件。

CI 在 Windows、macOS 和 Linux 上运行核心回归及两种分发构建。页面渲染环境应另行配置并完成逐页验收，构建通过不代表页面视觉已通过。新模板需登记来源、使用说明、指纹与结构契约，并通过同一分发检查。
