#!/bin/bash
# 拡張の開発用ディレクトリを現在のワークツリーに向ける
LINK="$HOME/outlook-ext-dev"
EXT_DIR="$(cd "$(dirname "$0")" && pwd)"

# シンボリックリンクを張り替え
rm -f "$LINK"
ln -s "$EXT_DIR" "$LINK"

echo "拡張ディレクトリを切り替えました:"
echo "  $LINK -> $EXT_DIR"

# Chrome に Cmd+Shift+U を送って拡張をリロード
osascript -e '
tell application "Google Chrome" to activate
delay 0.3
tell application "System Events"
    keystroke "u" using {command down, shift down}
end tell
'

echo "拡張をリロードしました"
