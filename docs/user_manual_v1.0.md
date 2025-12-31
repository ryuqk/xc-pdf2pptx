# PDF2PPTX Converter ユーザーマニュアル
**Version:** 1.0

## 1. はじめに
このツールは、PDFや画像ファイルを編集可能なPowerPoint (.pptx) ファイルに変換するツールです。AI技術 (Google Gemini) を使用して、文字の位置やフォント、太字などを高精度に認識し、元の見た目を保ったまま編集できるようにします。

## 2. インストールと準備

### 2.1. 必要なもの
- Google Gemini API キー
    - [Google AI Studio](https://aistudio.google.com/) から取得してください（無料枠あり）。

### 2.2. セットアップ (初回のみ)
1.  配布されたフォルダ (`PDF2PPTX`) を任意の場所に配置します。
2.  フォルダ内の `.env.example` ファイルをコピーし、名前を `.env` に変更します。
3.  `.env` ファイルをメモ帳などで開き、`GOOGLE_API_KEY=` の続きに、取得したAPIキーを貼り付けて保存します。
    ```text
    GOOGLE_API_KEY=AIzaSyBxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    ```

## 3. ツールの使い方

### 3.1. アプリの起動
フォルダ内の `PDF2PPTX_Converter.exe` をダブルクリックして起動します。

### 3.2. 画面構成と操作
画面は大きく3つのエリアに分かれています。

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

1.  **Settings (上部設定エリア)**
    *   **API Key**: Gemini APIキーを入力します。
        *   **Save**: 入力したAPIキーを保存します（`.env`ファイルに書き込まれます）。次回起動時から自動で入力されます。
        *   ※ Windowsのシステム環境変数 `GOOGLE_API_KEY` が設定されている場合は、それが優先的に表示されます。
    *   **Mode**: 変換モードを選択します。通常は `Standard` で、高精度を求める場合は `Text Focus` を選びます。
    *   **Font Scale (1.1x)**: チェックを入れると、AIが推定した文字サイズを1.1倍に少し大きくします。読みやすさを優先する場合に推奨します。

2.  **Files (中央ファイルリスト)**
    *   この枠内に変換したい **PDFファイル** や **画像ファイル** をドラッグ＆ドロップします。
    *   一度に複数のファイルを登録でき、上から順番に連続処理されます。
    *   **Add Files...**: ボタンからファイルを選択することもできます。
    *   **Clear List**: リストを空にします。

3.  **Output & Action (下部実行エリア)**
    *   **Output Folder**: 変換後のPPTXを保存するフォルダを指定します。空欄のままだと、元のファイルと同じ場所に保存されます。
    *   **変換開始 (Start Conversion)**: 青いボタンを押すと処理を開始します（APIキー必須）。
    *   **停止 (Cancel)**: 処理を中断します。即時停止ではありませんが、キリの良いタイミング（現在のページ完了時など）で止まります。
    *   **Logs**: 画面最下部に処理の進捗状況が表示されます。

### 3.3. 変換の手順
1.  APIキーが空欄の場合は入力し、**[Save]** を押しておくと便利です。
2.  PDFファイルまたは画像ファイル（PNG, JPGなど）をウィンドウ中央のリストに **ドラッグ＆ドロップ** します（複数可）。
3.  **[Start Conversion]** ボタンを押します。
4.  プログレスバーが進み、変換処理が始まります。
5.  完了するとダイアログが表示されます。
    *   **Complete**: すべて正常に完了しました。
    *   **Cancelled**: ユーザー操作により中断されました。

## 4. トラブルシューティング

### Q. エラーが出て変換できない
- **"Invalid API Key" と表示される**: APIキーが間違っているか、有効ではありません。設定画面で正しいキーが入力されているか確認し、**[Save]** してください。
- **その他のエラー**: インターネット接続やファイルが壊れていないか確認してください。
- **ファイルが壊れていませんか？** PDFビュワーなどでファイルが開けるか確認してください。

### Q. 文字の位置が少しずれる
- AIの認識精度による限界がありますが、**Mode** を `text_focus` にすることで、背景画像の上に文字を配置するため、見た目のズレは気になりにくくなります。

### Q. 変換が遅い
- ページ数が多い場合や、高解像度の画像の場合、AIの処理に時間がかかります。気長にお待ちください。
