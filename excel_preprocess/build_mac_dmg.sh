#!/bin/bash

echo "============================================"
echo "      Building GEpmTool for macOS (.app + .dmg)"
echo "============================================"

# 切到腳本所在位置
cd "$(dirname "$0")"

# 設定輸出目錄為腳本上一級目錄的 Tool_Pack_mac
OUTPUT_DIR="../Tool_Pack_mac"

# 若資料夾不存在則建立
if [ ! -d "$OUTPUT_DIR" ]; then
    echo "Creating output folder: $OUTPUT_DIR"
    mkdir -p "$OUTPUT_DIR"
fi

# 進入輸出資料夾
cd "$OUTPUT_DIR"

# 清除舊 build/dist
echo "Cleaning previous build..."
rm -rf build dist GEpmTool.spec GEpmTool.dmg

# 回到 excel_preprocess 取得原始腳本路徑
SCRIPT_ROOT="$(dirname "$(cd "$(dirname "$0")" && pwd)")/excel_preprocess"

echo "Running PyInstaller..."
pyinstaller \
    --distpath "$OUTPUT_DIR/dist" \
    --workpath "$OUTPUT_DIR/build" \
    --specpath "$OUTPUT_DIR" \
    --windowed \
    --name GEpmTool \
    --add-data "$SCRIPT_ROOT/ui_GEpmToolUI.py:." \
    --add-data "$(dirname "$SCRIPT_ROOT")/Doc/report_demo.xlsx:Doc" \
    --add-data "$(dirname "$SCRIPT_ROOT")/Doc/logo.png:Doc" \
    "$SCRIPT_ROOT/GUI_Tool.py"

APP_PATH="$OUTPUT_DIR/dist/GEpmTool.app"
DMG_PATH="../PM_Tool/Mac/GEpmTool.dmg"

# 檢查 .app 是否生成成功
if [ ! -d "$APP_PATH" ]; then
    echo "❌ Build failed: No .app created!"
    exit 1
fi

echo "============================================"
echo "   App build success! Creating DMG..."
echo "============================================"

# 如果已有舊的 DMG，先刪除以避免 hdiutil create 報錯
if [ -f "$DMG_PATH" ]; then
    rm -f "$DMG_PATH"
fi

# 建立 dmg
hdiutil create -volname "GEpmTool" -srcfolder "$APP_PATH" -format UDZO "$DMG_PATH"

echo "============================================"
echo "   🎉 Build Complete!"
echo "   App: $APP_PATH"
echo "   DMG: $DMG_PATH"
echo "============================================"