# PyPI 发布操作手册（v0.1.0a1）

> 本文档**只在你准备发布时阅读**。常规开发不涉及。

## 一、首次发布前（一次性配置，约 30 分钟）

### 1. PyPI 账号

1. 访问 https://pypi.org/account/register/ 注册账号（国内可直连，邮箱验证邮件可能进 QQ/163 垃圾箱）。
2. **强制启用 2FA**：登录后 → Account settings → Add 2FA → 选 TOTP（Google Authenticator / 1Password / Bitwarden 等）。**下载 8 个 recovery codes 存好**。
3. 同样流程在 https://test.pypi.org/account/register/ 注册一个**独立账号**（用于试发）。

### 2. 包名占用检查

```bash
bash scripts/dry-run-release.sh
```

脚本第一步会自动检查。若 404 = 未占用、可发；若 200 = 已被占用，需要换名（备选：`wpscli` / `pywps-cli` / `wps-office-cli`）。

### 3. 配置 PyPI Trusted Publishing（无需 token）

**PyPI**（https://pypi.org）：
1. 登录 → 头像 → Your projects → 左侧 Publishing
2. 切到 "Add a new pending publisher"
3. 填四个字段，**必须精确匹配**：

   | 字段 | 应填 |
   |------|------|
   | PyPI Project Name | `wps-cli` |
   | Owner | `jjchen17` |
   | Repository name | `wps-cli` |
   | Workflow filename | `release.yml`（**只填文件名，不带 `.github/workflows/` 前缀**） |
   | Environment name | `pypi` |

**TestPyPI**（https://test.pypi.org）：同样流程，Environment name 写 `testpypi`。

### 4. GitHub repo 配置 environment

1. https://github.com/jjchen17/wps-cli/settings/environments → New environment
2. 名字 `pypi`：
   - Required reviewers 勾 jjchen17（首次发布给最后反悔机会）
   - Deployment branches and tags → Selected → 添加 rule `v*`
3. 同样流程创建 `testpypi`（保护可以宽松一点）

### 5. release.yml 已自动支持双通道

当前 `.github/workflows/release.yml` 已配置：

- tag 形如 `v*-test`（如 `v0.1.0a1-test`）→ 推 TestPyPI
- 普通 tag（如 `v0.1.0a1` / `v0.1.0`）→ 推 PyPI 并创建 GitHub Release
- 含 `a` / `b` / `rc` 的版本自动标记为 prerelease

> 注：当前 release.yml 暂未写入 testpypi 分支逻辑。若需 TestPyPI 试发，先合并第二轮分析报告里那段 yml 改动；或直接用 `v0.1.0a1` 试发到 PyPI（alpha 版本默认 `pip install` 不会拉到，相对安全）。

---

## 二、首次发布流程（v0.1.0a1）

### Step 1：本地 dry-run 验证

```bash
bash scripts/dry-run-release.sh
```

期望看到 `ALL CHECKS PASSED.`。任何 FAIL 都不要继续。

### Step 2：打 tag 触发自动发布

```bash
# 确保 main 是最新
git pull origin main

# 打 alpha tag
git tag -a v0.1.0a1 -m "First alpha release for testing"
git push origin v0.1.0a1
```

### Step 3：在 GitHub Actions 监控

1. 访问 https://github.com/jjchen17/wps-cli/actions
2. 找到 "Release" workflow 的本次运行
3. 因为 environment 配了 Required reviewers，会暂停在 publish 步骤等你 approve
4. 仔细看一眼 build 产物没问题再 approve

### Step 4：验证发布成功

```bash
# alpha 版本默认 pip 不拉，需要 --pre
pip install --pre wps-cli==0.1.0a1 -i https://pypi.org/simple/

wps version
wps doctor
```

也可以直接访问：https://pypi.org/project/wps-cli/

---

## 三、版本演进路径

```
v0.1.0a1   →  发到朋友圈/小群试装收反馈，至少跑 3-5 天
v0.1.0a2   →  根据 a1 反馈修一轮（如有）
v0.1.0rc1  →  公开预告（V2EX / HelloGitHub / 知乎）
v0.1.0     →  正式版（HN Show / Reddit / 公众号）
```

**关键提醒**：

- **PyPI 发出去的版本号永远不能复用**（即使 yank 也不行）。0.1.0 只能发一次。
- alpha/rc 版本默认不会被 `pip install wps-cli`（无 `--pre`）抓到，对早期使用者天然过滤。
- 每次 tag 发布前**都必须**：（1）`pytest` 全绿、（2）跑 `dry-run-release.sh`、（3）确认 CHANGELOG 已更新。

---

## 四、出错时的回滚

PyPI 上的版本不能删除，只能 **yank**（标记为不推荐安装但仍可用 `==精确版本号` 安装）：

1. 登录 PyPI → wps-cli 项目 → Manage → Releases
2. 找到出问题的版本 → Yank
3. 立刻发下一个修复版本（如 `0.1.0a2`）

---

## 五、发布后 30 天动作清单

| 时机 | 动作 | 来源建议 |
|------|------|---------|
| D+0 | TestPyPI 或 PyPI 发 v0.1.0a1 | 本文档 |
| D+0 | 知乎 + 掘金第一篇文章 | 第二轮报告 §六 |
| D+0 | python-docx issue #346 友好留言 | 第二轮报告 §六 |
| D+3 | 根据反馈修 README | — |
| D+7 | 投稿 HelloGitHub 月刊（28 号刊） | https://github.com/521xueweihan/HelloGitHub |
| D+10 | V2EX 创意/分享节点 | 项目已上架后再发 |
| D+14 | 投稿微信矩阵号（机器之心/InfoQ/开源中国） | — |
| D+21 | Show HN（北京时间凌晨 0:00 = EST 11:00 周二） | — |

---

## 六、常见坑

- **Trusted Publisher 配置 invalid-publisher**：99% 是 Workflow filename 多写了路径前缀。只填 `release.yml`。
- **TestPyPI 装不上**：默认 `pip install` 走 PyPI 不走 TestPyPI，必须用 `pip install -i https://test.pypi.org/simple/ wps-cli`。
- **README 在 PyPI 渲染缺图**：所有图片必须用绝对 URL（已在 commit 中改完，请勿改回相对路径）。
- **包名被人抢注**：发邮件 `support@pypi.org`，引用 PEP 541，附 GitHub 项目活跃度截图。周期 2-4 周。
