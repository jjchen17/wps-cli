# 贡献指南

感谢你对 wps-cli 的关注。

## 开发环境

需求：

- Windows 10/11 (COM 自动化测试需要)
- Python 3.10+
- WPS Office 2019+

安装：

```bash
git clone https://github.com/jjchen17/wps-cli.git
cd wps-cli
pip install -e ".[dev]" -i https://pypi.tuna.tsinghua.edu.cn/simple
pre-commit install
```

## 工作流

1. 在 [Issues](https://github.com/jjchen17/wps-cli/issues) 中确认或创建议题。
2. 从 `main` 分支拉新分支，命名建议 `feat/xxx`、`fix/xxx`、`docs/xxx`。
3. 写代码 + 写测试（80% 覆盖率门槛）。
4. 本地跑通：

   ```bash
   ruff check src/ tests/
   ruff format --check src/ tests/
   pytest --cov=src --cov-report=term-missing
   ```

5. 提交 PR，标题用 [Conventional Commits](https://www.conventionalcommits.org/zh-hans/) 规范：
   - `feat: ...` 新功能
   - `fix: ...` Bug 修复
   - `refactor: ...` 重构
   - `docs: ...` 文档
   - `test: ...` 测试
   - `chore: ...` 构建/工具链

## 代码规范

- **类型注解**：所有公开函数必须有完整类型注解。
- **不变性**：优先使用 `@dataclass(frozen=True)`、`NamedTuple`、不可变容器。
- **小文件**：单文件 200–400 行最佳，超过 800 行考虑拆分。
- **错误处理**：抛 `WpsCliError` 子类，附带 `suggestion` 与 `context`，避免裸 `except Exception`。
- **路径输入**：CLI 层必须经过 `ensure_safe_input_path` / `ensure_safe_output_path` 校验，禁止裸 `Path(arg)` 透传到 Service 层。
- **COM 调用**：所有 `Documents.Open` / `Workbooks.Open` / `Presentations.Open` 必须显式传 `ReadOnly`、`AddToRecentFiles=False` 等安全参数。

## 测试

- 单元测试用 `MockComBackend`，无需真实 WPS 环境。
- 集成测试需要本地有 WPS 安装；CI 在 `windows-latest` 上跑。
- 涉及 COM 行为的测试请加 `@pytest.mark.integration` 标记。
- 新增公开 API 必须配对单元测试。

## 安全相关贡献

如果你修复的是安全相关问题（路径穿越、注入、宏执行等），请：

1. **不要在公开 issue 中讨论攻击细节**。
2. 直接邮件 `948881912@qq.com` 联系维护者。
3. PR 描述中说明威胁模型与修复方式。
4. 加测试覆盖恶意输入。

## 提交检查清单

- [ ] `pytest` 全部通过
- [ ] `ruff check` 无报错
- [ ] `ruff format` 已应用
- [ ] 新增/修改公开 API 已更新文档
- [ ] CHANGELOG.md 添加 Unreleased 条目
- [ ] 没有提交 `.claude/`、虚拟环境、IDE 配置等本地文件

## 许可证

提交即表示你的贡献以 [MIT 许可证](LICENSE) 释出。
