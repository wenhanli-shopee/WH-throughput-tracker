# WH Throughput Tracker — Auto-Update Setup

Dashboard URL (after setup): `https://liwenhanmia.github.io/WH-throughput-tracker/`

Last updated by GitHub Actions every Monday 08:00 SGT.

---

## 一次性配置步骤 (One-time Setup)

### Step 1 — 创建 GitHub 仓库并开启 GitHub Pages

1. 登录 GitHub，进入 `https://github.com/liwenhanmia/WH-throughput-tracker`
2. 如果仓库还不存在，点 **New repository**，名字填 `WH-throughput-tracker`，选 **Public**
3. 进入仓库 → **Settings** → **Pages**
4. Source 选 **Deploy from a branch**
5. Branch 选 **main**，Folder 选 **/ (docs)**，点 Save
6. 等 1-2 分钟，页面会显示你的 URL：`https://liwenhanmia.github.io/WH-throughput-tracker/`

---

### Step 2 — 创建 Google Service Account

1. 打开 [Google Cloud Console](https://console.cloud.google.com/)
2. 创建一个新项目，名字随意，比如 `wh-tracker`
3. 左侧菜单 → **APIs & Services** → **Enabled APIs** → 点 **+ ENABLE APIS AND SERVICES**
4. 搜索 **Google Sheets API**，点进去 → **Enable**
5. 左侧菜单 → **APIs & Services** → **Credentials**
6. 点 **+ CREATE CREDENTIALS** → **Service account**
7. 填名字（比如 `wh-tracker-reader`），其他默认，点 **Done**
8. 点刚创建的 Service Account → **Keys** tab → **ADD KEY** → **Create new key** → JSON → **Create**
9. 浏览器会自动下载一个 `.json` 文件，**妥善保管，不要上传到 GitHub**

---

### Step 3 — 把 Service Account 加为 GSheet 的 Viewer

打开下载的 JSON 文件，找到 `"client_email"` 字段，复制那个邮箱地址，格式类似：
```
wh-tracker-reader@wh-tracker-xxxxx.iam.gserviceaccount.com
```

然后对你的**三个 GSheet 文件**（PHB、PHL、PHIXC）分别操作：
1. 打开 GSheet → 右上角 **Share**
2. 把上面那个邮箱粘贴进去
3. 权限选 **Viewer**
4. 点 **Send**（不需要通知，可以忽略"notify"）

> ⚠️ 如果你的 GSheet 里有 `IMPORTRANGE` 引用其他文件，那些源文件也需要同样操作。

---

### Step 4 — 把 Dashboard HTML 放入仓库

1. 把当前的 `PHB_WH_Capacity_Dashboard.html` 重命名为 `index.html`
2. 放到仓库的 `docs/` 目录下

这是数据注入的模板，脚本会直接修改这个文件里的 `const WH={...}` 和 `const BM_PHB={...}` 等数据块。

---

### Step 5 — 把 Secrets 存入 GitHub

进入仓库 → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**

需要添加 **4 个 Secrets**：

| Secret 名称 | 内容 |
|-------------|------|
| `GOOGLE_SERVICE_ACCOUNT_JSON` | 把第 Step 2 下载的整个 JSON 文件内容粘贴进来（全部文字） |
| `SHEET_ID_PHB` | PHB 的 GSheet ID（URL 里 `/d/` 和 `/edit` 之间那串字符） |
| `SHEET_ID_PHL` | PHL 的 GSheet ID |
| `SHEET_ID_PHIXC` | PHIXC 的 GSheet ID |

**如何找 Sheet ID：**
```
https://docs.google.com/spreadsheets/d/ 【这里就是 Sheet ID】 /edit
```

---

### Step 6 — 上传所有文件到仓库

仓库结构应该是：
```
WH-throughput-tracker/
├── .github/
│   └── workflows/
│       └── update.yml
├── scripts/
│   └── fetch_and_build.py
├── docs/
│   └── index.html          ← 你的 dashboard HTML（改名后放这里）
├── requirements.txt
└── README.md
```

用 Git 上传：
```bash
git clone https://github.com/liwenhanmia/WH-throughput-tracker.git
cd WH-throughput-tracker
# 把上面所有文件放进来
git add .
git commit -m "Initial setup"
git push
```

---

### Step 7 — 手动触发一次，验证是否正常

1. 进入仓库 → **Actions** tab
2. 左侧选 **Weekly Dashboard Update**
3. 右侧点 **Run workflow** → **Run workflow**
4. 等待约 30-60 秒
5. 如果绿色 ✅ = 成功；红色 ❌ = 点进去看错误日志

---

## 之后每周自动运行

配置好之后，每周一 08:00 新加坡时间自动：
1. 从三个 GSheet 读取最新数据
2. 重新生成 `docs/index.html`
3. 自动 commit 并 push
4. GitHub Pages 在 1-2 分钟内更新

你也可以随时手动触发（Step 7 的方法）。

---

## 常见问题

**Q: Actions 报错 `RESOURCE_EXHAUSTED` 或 `PERMISSION_DENIED`**
A: Service Account 没有被加为 GSheet 的 Viewer，重新检查 Step 3。

**Q: `WorksheetNotFound` 错误**
A: GSheet 的 tab 名称和脚本里的不一致。打开 `scripts/fetch_and_build.py`，
找 `worksheet('IB Model')` 这样的地方，改成你实际的 tab 名称。

**Q: HTML 更新了但数据还是旧的**
A: 检查 `fetch_and_build.py` 里 `read_actual_data()` 函数，
需要根据你 GSheet 的实际行列位置调整读取逻辑。

**Q: 想立刻更新，不等周一**
A: Actions → Weekly Dashboard Update → Run workflow（手动触发）。
