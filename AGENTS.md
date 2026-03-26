
## Non-Engineer Guardrails

- 破壊的操作（`rm -rf`, `git reset --hard`, 履歴改変）はユーザーの明示許可なしで実行しない
- 外部送信（メール/Slack投稿/本番デプロイ/公開URL共有）は実行前に必ず確認する
- `credentials*.json`, `.env`, token, secret をGitにcommitしない
- 高リスク変更（DBマイグレーション、権限設定変更、課金リソース作成）は実行前に理由と影響を提示する
- 不明点がある場合は推測で実行せず、選択肢と推奨案を提示して確認する
