# Render 免费部署指南（零服务器、零运维）

部署包已就绪，共 5 个文件：`quote_engine.py`、`wecom_bot.py`、`requirements.txt`、`Procfile`、`天航空派同行代理价11.7.xlsx`

## 第一步：把代码放上 GitHub（5分钟）

1. 注册/登录 [github.com](https://github.com)，点右上角 **+ → New repository**
2. 仓库名填 `quote-bot`，选 **Private**（私有，保护报价数据），点 Create
3. 页面上点 **uploading an existing file**，把这 5 个文件全部拖进去（不要放进子文件夹）
4. 点 **Commit changes**

## 第二步：部署到 Render（5分钟）

1. 注册/登录 [render.com](https://render.com)（可用 GitHub 账号直接登录）
2. 控制台点 **New + → Web Service**，授权连接你刚建的 `quote-bot` 仓库
3. 配置页面照抄：
   - **Name**: `quote-bot`
   - **Region**: Singapore（离中国最近）
   - **Branch**: main
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn wecom_bot:app --bind 0.0.0.0:$PORT --timeout 120`
4. 点 **Environment** 标签 → Add Environment Variable，添加 3 个：
   - `WECOM_TOKEN` = 你在企微后台设置的 Token
   - `WECOM_AESKEY` = 企微后台生成的 EncodingAESKey（43位）
   - `WECOM_CORPID` = 企业ID
5. 点 **Create Web Service**，等 2~3 分钟构建，状态变 **Live** 即成功
6. 页面顶部会显示你的专属地址，形如 `https://quote-bot-xxxx.onrender.com`，记下来

## 第三步：企微后台配置（5分钟）

1. 企业微信管理后台 → 应用管理 → 自建 → 创建应用「报价助手」
2. 进入应用 → 接收消息 → 设置 API 接收：
   - URL：`https://quote-bot-xxxx.onrender.com/wecom/callback`（换成你的地址）
   - Token / EncodingAESKey：**自己点随机生成**，并把这两个值同步填到 Render 的环境变量里（第二步第4点）
3. 点保存 → 企微立即验证，通过即接通
4. 回到应用详情 → 企业可信IP：填 Render 给的 IP（或开启"不校验"，测试阶段）
5. 在内部报价群添加机器人 / 或让销售在群内使用该应用，@机器人 发：
   `量勺7件套 39*28*25cm 17kg/箱 2箱 76117`
   → 机器人秒回报价

## 注意事项

- 免费版 **15分钟无请求会休眠**，下次调用唤醒需 5~20 秒；想常驻可升级 $7/月（仍比服务器便宜）
- **更新价格表**：在 GitHub 仓库里重新上传新的 xlsx（覆盖同名文件），Render 会自动重新部署，无需其他操作
- 仓库是 Private 的，报价数据不会公开
- 每次价格表更新也可以直接发我，我帮你改好引擎规则后给你新的部署包
