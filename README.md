# PDF2PPTX Converter
Google Gemini APIを活用して、PDFファイルや画像ファイルを編集可能なPowerPoint (.pptx) ファイルに変換するツールです。
レイアウト解析にAIを使用することで、これまで難しかった「元のレイアウトを保ったままのテキスト化」を高精度に実現します。

<svg xmlns="http://www.w3.org/2000/svg" width="800" height="600" viewBox="0 0 800 600">
  <!-- Main Window -->
  <rect x="0" y="0" width="800" height="600" fill="#f0f0f0" stroke="#333" stroke-width="2"/>
  <rect x="0" y="0" width="800" height="30" fill="#ffffff" stroke="#333" stroke-width="1"/>
  <text x="10" y="20" font-family="Segoe UI, sans-serif" font-size="14" fill="#333">PDF/Image to PPTX Converter (Gemini Powered)</text>
  <rect x="770" y="5" width="20" height="20" fill="none" stroke="#333" stroke-width="1"/>
  <line x1="772" y1="7" x2="788" y2="23" stroke="#333" stroke-width="1"/>
  <line x1="788" y1="7" x2="772" y2="23" stroke="#333" stroke-width="1"/>
  <!-- Settings Frame -->
  <rect x="10" y="40" width="780" height="60" rx="5" ry="5" fill="none" stroke="#ccc" stroke-width="1"/>
  <text x="20" y="50" font-family="Segoe UI" font-size="12" fill="#0078D7" font-weight="bold" bgcolor="#f0f0f0"> Settings </text>
  <!-- API Key -->
  <text x="30" y="80" font-family="Segoe UI" font-size="12">API Key:</text>
  <rect x="80" y="65" width="200" height="25" fill="white" stroke="#999"/>
  <text x="85" y="82" font-family="monospace" font-size="12">****************</text>
  <rect x="290" y="65" width="50" height="25" rx="3" fill="#e1e1e1" stroke="#999"/>
  <text x="300" y="82" font-family="Segoe UI" font-size="12">Save</text>
  <!-- Mode -->
  <text x="360" y="80" font-family="Segoe UI" font-size="12">Mode:</text>
  <rect x="400" y="65" width="100" height="25" fill="white" stroke="#999"/>
  <text x="405" y="82" font-family="Segoe UI" font-size="12">text_focus</text>
  <polygon points="485,73 495,73 490,80" fill="#333"/>
  <!-- Font Scale -->
  <rect x="520" y="70" width="14" height="14" fill="white" stroke="#333"/>
  <path d="M522 75 L526 79 L532 71" stroke="black" stroke-width="2" fill="none"/>
  <text x="540" y="82" font-family="Segoe UI" font-size="12">Font Scale (1.1x)</text>
  <!-- Files List Frame -->
  <rect x="10" y="110" width="780" height="300" rx="5" ry="5" fill="none" stroke="#ccc" stroke-width="1"/>
  <text x="20" y="120" font-family="Segoe UI" font-size="12" fill="#0078D7" font-weight="bold"> Files (Drag & Drop here) </text>
  <rect x="20" y="130" width="740" height="230" fill="white" stroke="#999"/>
  <text x="25" y="150" font-family="Segoe UI" font-size="12">J:\XBP\PDF2PPTX\input.pdf</text>
  <text x="25" y="170" font-family="Segoe UI" font-size="12">J:\XBP\PDF2PPTX\image.png</text>
  <!-- Scrollbar -->
  <rect x="760" y="130" width="20" height="230" fill="#f0f0f0" stroke="#999"/>
  <rect x="765" y="135" width="10" height="50" rx="5" fill="#ccc"/>
  <!-- List Buttons -->
  <rect x="20" y="370" width="80" height="25" rx="3" fill="#e1e1e1" stroke="#999"/>
  <text x="25" y="387" font-family="Segoe UI" font-size="12">Add Files...</text>
  <rect x="110" y="370" width="80" height="25" rx="3" fill="#e1e1e1" stroke="#999"/>
  <text x="120" y="387" font-family="Segoe UI" font-size="12">Clear List</text>
  <!-- Output & Action Frame -->
  <rect x="10" y="420" width="780" height="60" rx="5" ry="5" fill="none" stroke="none"/> <!-- Invisible container for logical grouping -->
  <text x="30" y="450" font-family="Segoe UI" font-size="12">Output Folder:</text>
  <rect x="115" y="435" width="200" height="25" fill="white" stroke="#999"/>
  <rect x="325" y="435" width="70" height="25" rx="3" fill="#e1e1e1" stroke="#999"/>
  <text x="335" y="452" font-family="Segoe UI" font-size="12">Browse...</text>
  <!-- Cancel Button -->
  <rect x="580" y="430" width="80" height="40" rx="3" fill="#FFA500" stroke="#cc8400"/>
  <text x="595" y="455" font-family="Meiryo UI" font-size="12" font-weight="bold" fill="black">Cancel</text>
  <!-- Start Button -->
  <rect x="670" y="430" width="120" height="40" rx="3" fill="#0078D7" stroke="#005a9e"/>
  <text x="680" y="455" font-family="Meiryo UI" font-size="12" font-weight="bold" fill="white">Start Conversion</text>
  <!-- Progress Bar -->
  <rect x="10" y="490" width="780" height="20" fill="#e0e0e0" stroke="#999"/>
  <rect x="10" y="490" width="300" height="20" fill="#00cc00" stroke="none"/> <!-- 30% progress -->
  <!-- Logs Frame -->
  <rect x="10" y="520" width="780" height="70" rx="5" ry="5" fill="none" stroke="#ccc" stroke-width="1"/>
  <text x="20" y="530" font-family="Segoe UI" font-size="12" fill="#0078D7" font-weight="bold"> Logs </text>
  <rect x="20" y="535" width="760" height="50" fill="white" stroke="#999"/>
  <text x="25" y="550" font-family="Consolas" font-size="10" fill="#333">Processing File 1/2: input.pdf</text>
  <text x="25" y="565" font-family="Consolas" font-size="10" fill="#333">  - Page 1/3...</text>
</svg>


## 主な機能

*   **高精度なレイアウト復元**: Gemini 3.0 Flash (またはその他のGeminiモデル) を使用し、テキストブロックと画像領域を識別してスライド上に再配置します。
*   **テキスト編集可能**: OCR結果をテキストボックスとして配置するため、変換後にPowerPoint上で自由に編集できます。
*   **画像/図表の保持**: テキスト以外の図や写真は画像としてスライドに配置されます。
*   **選べる2つのモード**:
    *   `Standard`: テキストと画像を個別に配置する標準モード。
    *   `Text Focus`: 背景を画像として敷き、その上に透明なテキストボックスを配置する、見た目の再現性を重視したモード。
*   **簡単操作**: ドラッグ＆ドロップ対応のGUIアプリが付属しています。

## 必要な環境

*   Python 3.10 以上推奨
*   Google Gemini API Key

## インストール手順

1.  リポジトリをクローンまたはダウンロードします。
2.  必要なライブラリをインストールします。

```bash
pip install -r requirements.txt
```

3.  環境変数を設定します。
    - プロジェクトルートに `.env.example` があるので、それをコピーして `.env` という名前に変更します。
    - `.env` ファイル内の `GOOGLE_API_KEY` にご自身のGemini APIキーを記述してください。

```text
GOOGLE_API_KEY=your_api_key_here
```

## 使い方

### GUIアプリを使用する場合 (推奨)

以下のコマンドでGUIアプリを起動します。

```bash
python gui_app.py
```

1.  **API Key設定**: 初回起動時はSettingsエリアにAPI Explorer等で取得したAPIキーを入力し「Save」を押してください。
2.  **ファイル追加**: 変換したいPDFや画像を画面中央のリストにドラッグ＆ドロップします。
3.  **変換開始**: 「Start Conversion」ボタンを押すと変換が始まります。

### コマンドライン (CLI) を使用する場合

バッチ処理などを行いたい場合は、コマンドラインから直接実行することも可能です。

```bash
python pdf2pptx.py input.pdf output.pptx --api_key "YOUR_API_KEY"
```

オプション:
*   `--mode`: `standard` か `text_focus` を指定 (デフォルト: standard)
*   `--font_scale`: フォントサイズの拡大縮小率 (デフォルト: 1.1)

## ドキュメント

詳細な仕様や操作マニュアルについては `docs` フォルダをご確認ください。

*   [ユーザーマニュアル (docs/user_manual_v1.0.md)](docs/user_manual_v1.0.md)

## ライセンス

MIT License (または適切なライセンスをここに記載)
