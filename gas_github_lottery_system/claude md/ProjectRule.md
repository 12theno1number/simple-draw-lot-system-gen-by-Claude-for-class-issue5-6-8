# 專案規則（Project Rules）

專案：抽獎系統 2.0（SDLC GitHub Lottery System）
建立日期：2026-10-19

## 1. Markdown 文件位置
- 任何新增的 `.md` 檔案，一律寫在 `claude md/` 資料夾內。
  - 例：`claude md/ProjectRule.md`、`claude md/IMPLEMENTATION_PLAN.md`

## 2. 程式檔案位置
- 任何用於測試與執行的程式檔案，一律寫在 `claude script/` 資料夾內：
  - 測試用程式 → `claude script/test/`
  - 專案執行用程式 → `claude script/project/`

## 資料夾結構

```
gas_github_lottery_system/
├── claude md/            # 所有新增的 .md 文件
├── claude script/
│   ├── project/          # 專案執行程式
│   └── test/             # 測試程式
└── ...
```
