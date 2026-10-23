# JMeter API テスト JMX 生成ツール（jmeter_api_tool）

REST API の定義（URL・HTTP メソッド・クエリパラメータ・ヘッダ・リクエストボディ）を **Excel または CSV** に記入するか、**OpenAPI 定義（JSON）** を指定すると、
**Apache JMeter 5.6.3** 用の JMX ファイルを自動生成し、CLI モードで実行して JTL・HTML レポート・ログを出力するツールです。

- RHEL 9.8 用: `jmeter_api_tool.sh`（bash + OS 標準の python3。追加インストール不要）
- Windows 11 用: `jmeter_api_tool.ps1` ＋ 起動用 `jmeter_api_tool.bat`（Windows PowerShell 5.1 / PowerShell 7。追加モジュール・Excel 不要）
- 両スクリプトの処理内容は同一で、同じ入力から同じ JMX・Excel・Markdown を出力します。

詳しい使い方は **`docs/JMeter_Usage_Guide.xlsx`**（Meiryo UI・モノトーン）または **`docs/JMeter_Usage_Guide.md`** を参照してください。

## フォルダ構成

| パス | 内容 |
|---|---|
| `jmeter_api_tool.sh` | RHEL 9.8 用スクリプト |
| `jmeter_api_tool.ps1` / `jmeter_api_tool.bat` | Windows 11 用スクリプト / 起動ファイル（実行ポリシーを回避して .ps1 を起動） |
| `input/api_definition.xlsx` | 入力テンプレート（Excel）：「API一覧」「設定」「記入方法」シート |
| `input/api_definition.csv` / `input/api_definition_settings.csv` | 入力テンプレート（CSV）：API 一覧 / 設定 |
| `output/api_test.jmx` | テンプレートの記入例から生成した JMX（サンプル） |
| `output/data/testdata.csv` / `testdata.xlsx` | CSV / Excel データ読込のサンプルデータ |
| `docs/JMeter_Usage_Guide.xlsx` / `.md` | 利用方法（manual コマンドで再出力できます） |

## クイックスタート

### RHEL 9.8

```bash
chmod +x jmeter_api_tool.sh
./jmeter_api_tool.sh init                                   # 入力テンプレートを作成（input/）
./jmeter_api_tool.sh generate -i input/api_definition.xlsx  # JMX を生成（output/）
./jmeter_api_tool.sh generate -i openapi.json --seed 12345  # OpenAPI 定義（JSON）から JMX を生成
./jmeter_api_tool.sh run -t output/api_test.jmx             # CLI 実行（result/JMX名_日時/）
./jmeter_api_tool.sh run -t output/api_test.jmx --threads 10 --loops 100   # 大量実行
./jmeter_api_tool.sh manual                                 # 利用方法を出力（docs/）
```

### Windows 11

```powershell
.\jmeter_api_tool.bat init
.\jmeter_api_tool.bat generate -i input\api_definition.xlsx
.\jmeter_api_tool.bat generate -i openapi.json -Seed 12345
.\jmeter_api_tool.bat run -t output\api_test.jmx
.\jmeter_api_tool.bat run -t output\api_test.jmx -Threads 10 -Loops 100
.\jmeter_api_tool.bat manual
```

`run` コマンドには Java（17 推奨）と JMeter 5.6.3 が必要です（`JMETER_HOME` を設定するか `--jmeter-home` / `-JMeterHome` で指定）。
`init` / `generate` / `manual` は JMeter・Java が無くても動作します。

## OpenAPI 定義（JSON）からの生成

`-i` に OpenAPI 3.0.x の JSON（拡張子 `.json`）を指定すると、Excel / CSV と同じ「API一覧」の 1 行 = 1 API に変換してから JMX を生成します。
定義の読み込みとテストデータの生成は JMeter_Tools の `openapi2jmx`（`OpenApi2Jmx.ps1`）と同じ処理で、同じシードなら同じテストデータになります。

| API一覧の列 | OpenAPI から設定する値 |
|---|---|
| No / 有効 / API名 | 定義の順番 / Y / `メソッド パス [operationId]`（ANY は `ANY(GET) パス`） |
| URL | パス（パスパラメータはデータ型に合ったランダム値）。接続先は設定 `base_url` |
| クエリパラメータ・ヘッダ・ボディ形式・リクエストボディ | パラメータ・requestBody の定義から生成（style / explode、JSON / フォーム / multipart / XML / テキスト） |
| 期待ステータス | 2xx の応答コード（`2XX` だけの場合は空欄） |
| 備考 | 概要・説明・operationId・認可（security）・API Gateway 統合など |

- 設定は Excel / CSV と同じく `-s` または「入力ファイル名_settings.csv」から読み込みます。設定ファイルで空欄の `base_url`（`servers[0].url`。無い場合は `http://localhost:8080`）、`test_plan_name`、`output_file`（`<タイトル>.jmx`）は OpenAPI 定義の値を使います。
- 認証ヘッダ（Authorization・X-API-KEY）はランダム値では通らないため生成しません。設定 `common_headers` に指定してください。
- OpenAPI 用のオプション（sh / PowerShell）: `--seed` / `-Seed`（乱数シード）、`--any-method` / `-AnyMethod`、`--required-only` / `-RequiredOnly`、`--use-examples` / `-UseExamples`、`--placeholder-value` / `-PlaceholderValue`（意味は JMeter_Tools と同じ）。
- GET / DELETE のリクエストボディは Excel / CSV と同じ規則で送信しません（警告を表示）。YAML・Swagger 2.0・外部ファイルへの `$ref` は対象外です。

## 生成される JMX の主な内容

- 既定は 1 スレッド × 1 回。`-JTHREADS` `-JLOOPS` `-JRAMP_UP`（`-JDURATION`）で大量件数の負荷試験に切り替え
- 接続先は `${__P(HOST,...)}` 形式の共通変数（`-JHOST=` で環境切替）
- API ごと: ヘッダ、定数タイマ（待機）、前処理・後処理（JSR223 / Groovy）、正規表現抽出、JSON 抽出、ステータス・本文・応答時間の検証
- 共通: 連番（Counter）、乱数、CSV データ読込、Excel データ読込（Apache POI 不要）、クッキー管理、共通前処理・後処理、スループット制御（既定は無効）
- 結果出力: JTL（CSV）、エラー詳細（XML）、結果ログ（api_result.log）、jmeter.log、HTML レポート（run コマンド）
- 要素名は `[必須]` / `[任意]` で始まり、`[任意]` の要素は JMeter GUI で削除しても他の要素に影響しません。
  生成時に設定で `OFF`（生成しない）/ `DISABLE`（無効状態で生成）も選べます。

## 動作確認済みの環境

| 環境 | 確認内容 |
|---|---|
| RHEL 9.8（Red Hat UBI 9.8 コンテナ、bash 5.1.8、python 3.9.25、OpenJDK 17.0.20、ロケール未設定） | init / generate / manual / run（JMeter 5.6.3 で実行、HTML レポート生成まで） |
| Windows 11（Windows PowerShell 5.1、Java 17.0.11） | init / generate / manual / run（成功・失敗・`-FailOnError`） |
| PowerShell 7（公式コンテナ） | init / generate / manual |
| Microsoft Excel（Office 16） | 生成した xlsx の読込（修復なし）、全セルの折り返し高さ（はみ出しなし）、Excel で編集・保存したテンプレートの再読込 |

- RHEL 版と Windows 版の出力（JMX・xlsx の各 XML・md・csv）が、生成日時以外は完全に一致することを確認しています。
- 生成した JMX は JMeter 5.6.3 の CLI / GUI で読み込みエラーがないこと、記入例の 4 API がエラー 0 件で実行できることを確認しています
  （記入例は公開サービス postman-echo.com を呼び出すため、インターネット接続が必要です）。

## 注意事項

- RHEL へは改行コード LF のまま転送してください（CRLF になった場合: `sed -i 's/\r$//' jmeter_api_tool.sh`）。
- `.ps1` は UTF-8（BOM 付き）で保存しています。編集する場合も BOM 付き UTF-8 のまま保存してください（Windows PowerShell 5.1 の文字化け防止）。
- 両スクリプトには同じ定義データ（設定項目・テンプレート・Groovy スクリプト・利用方法の本文）が埋め込まれています。改修する場合は両方を同じ内容に変更してください。
