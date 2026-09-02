# 武汉中心仓质量数据看板

在线地址：**https://zffzms1995.github.io/wuhan-quality-dashboard/**

在线看板 + 一键更新脚本。数据来源：飞书云表格「📊26年X月武汉中心仓质量数据」（每月一张表）+「X月提升动作留底表」（错题/培训数据）。看板顶部可按月份切换查看历史月份。

## 每天更新看板（只需一步）

1. 在飞书云表格里录入当天质检数据（和以前一样）
2. 打开「终端」（Terminal），粘贴运行：

```bash
cd ~/Desktop/仓内看板 && python3 update_dashboard.py
```

脚本会自动：从飞书导出最新表格 → 解析 → 生成数据文件和主板图片 → 归档当月数据（供看板月份切换回看）→ 提交到 GitHub（页面自动更新，1-2 分钟后生效）。

看到 `推送成功` 就完成了。

**换月时**：在飞书新建当月质量数据表和留底表并分享给应用，然后把新表链接发给 Claude 更新 token（或直接改脚本顶部 `DEFAULT_FILE_TOKEN` / `CUOTI_FILE_TOKEN`，或运行命令时加 `--token 新token --cuoti-token 新留底token`）。

## 本地预览（可选）

```bash
cd ~/Desktop/仓内看板 && python3 -m http.server 8765
```

然后浏览器打开 http://localhost:8765/

## 常见问题

- **看板上的数据不对**：看板是只读展示，不能在上面直接改。回到飞书原表格（质量数据表或错题集）修改，改完重新运行一次更新命令即可
- **运行报错说缺 openpyxl**：先装一次 `python3 -m pip install --user openpyxl`
- **提示数据加载失败**：多半是双击 html 打开了（file:// 方式读不了数据），用上面的本地预览命令
- **脚本中途报错**：不会覆盖旧数据，把错误信息发给 Claude 处理
- **主板图片没显示**：图片随数据文件一起推送，等 1-2 分钟刷新即可

## 文件说明

| 文件 | 作用 |
|---|---|
| `update_dashboard.py` | 一键更新脚本（导出→解析→生成→推送） |
| `index.html` | 看板页面 |
| `dashboard_data.json` | 看板数据（脚本生成，勿手改） |
| `trend_data.json` | 趋势数据（脚本生成，勿手改） |
| `images/` | 主板审核图片（脚本生成） |
| `archive/YYYY-MM/` | 各月归档数据+图片（看板月份切换用，脚本自动维护） |
| `months.json` | 月份列表（看板月份下拉用，脚本生成） |
