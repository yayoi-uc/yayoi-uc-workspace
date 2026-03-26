#!/bin/bash
# Git Auto-Sync: ファイル変更を監視して自動commit & push
# Usage: ./git-auto-sync.sh [--daemon]

REPO_DIR="${WORKSPACE_DIR:-${HOME}/clawd}"
LOG_FILE="${REPO_DIR}/logs/git-auto-sync.log"
LOCK_FILE="/tmp/git-auto-sync.lock"
DEBOUNCE_SEC=5  # 変更後の待機時間（連続変更をまとめる）

# ログ出力
log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $1" | tee -a "$LOG_FILE"
}

# 監視方式を判定
detect_watcher() {
    if command -v inotifywait >/dev/null 2>&1; then
        echo "inotify"
    elif command -v fswatch >/dev/null 2>&1; then
        echo "fswatch"
    else
        echo "polling"
    fi
}

# 既に実行中かチェック
if [ -f "$LOCK_FILE" ]; then
    PID=$(cat "$LOCK_FILE")
    if kill -0 "$PID" 2>/dev/null; then
        echo "Already running (PID: $PID)"
        exit 1
    fi
fi

# デーモンモード
if [ "$1" = "--daemon" ]; then
    nohup "$0" >> "$LOG_FILE" 2>&1 &
    echo $! > "$LOCK_FILE"
    echo "Started git-auto-sync daemon (PID: $!)"
    exit 0
fi

# PIDファイル作成
echo $$ > "$LOCK_FILE"
trap "rm -f $LOCK_FILE" EXIT

cd "$REPO_DIR" || exit 1
mkdir -p "$(dirname "$LOG_FILE")"
log "Git auto-sync started. Watching: $REPO_DIR"

# 監視対象外のパターン
EXCLUDE="\.git|node_modules|\.swp|\.tmp|__pycache__|\.pyc|logs/git-auto-sync"

# 同期実行
do_sync() {
    cd "$REPO_DIR" || return
    
    # 変更があるかチェック
    if [ -z "$(git status --porcelain)" ]; then
        return
    fi
    
    # 変更されたファイルを取得
    CHANGED=$(git status --porcelain | head -5 | awk '{print $2}' | tr '\n' ', ' | sed 's/,$//')
    
    git add -A
    git commit -m "auto-sync: ${CHANGED}"
    
    if git push 2>&1; then
        log "Pushed: ${CHANGED}"
    else
        log "Push failed: ${CHANGED}"
    fi
}

WATCHER=$(detect_watcher)
log "Watcher mode: $WATCHER"

# メインループ
LAST_SYNC=0
while true; do
    CHANGE=""
    case "$WATCHER" in
        inotify)
            CHANGE=$(inotifywait -r -q -t 60 \
                --exclude "$EXCLUDE" \
                -e modify,create,delete,move \
                "$REPO_DIR" 2>/dev/null)
            ;;
        fswatch)
            CHANGE=$(fswatch -1 --exclude "$EXCLUDE" "$REPO_DIR" 2>/dev/null)
            ;;
        *)
            sleep 10
            if [ -n "$(git status --porcelain)" ]; then
                CHANGE="polling-detected"
            fi
            ;;
    esac

    if [ -n "$CHANGE" ]; then
        # デバウンス: 連続変更をまとめる
        NOW=$(date +%s)
        if [ $((NOW - LAST_SYNC)) -ge $DEBOUNCE_SEC ]; then
            sleep $DEBOUNCE_SEC  # 追加の変更を待つ
            do_sync
            LAST_SYNC=$(date +%s)
        fi
    fi
done
