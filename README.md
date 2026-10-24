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
| `input/openapi_config.csv` | OpenAPI 用設定ファイルのテンプレート：接続先（PROTOCOL / HOST / PORT）・HTTP ヘッダ・アクセストークンのフォルダ |
| `output/api_test.jmx` | テンプレートの記入例から生成した JMX（サンプル） |
| `output/data/testdata.csv` / `testdata.xlsx` | CSV / Excel データ読込のサンプルデータ |
| `docs/JMeter_Usage_Guide.xlsx` / `.md` | 利用方法（manual コマンドで再出力できます） |

## クイックスタート

### RHEL 9.8

```bash
chmod +x jmeter_api_tool.sh
./jmeter_api_tool.sh init                                   # 入力テンプレートを作成（input/）
./jmeter_api_tool.sh generate -i input/api_definition.xlsx  # JMX を生成（output/）
./jmeter_api_tool.sh generate -i input/api_definition.xlsx -n order_test.jmx   # 出力する JMX のファイル名を指定
./jmeter_api_tool.sh generate -i openapi.json --seed 12345  # OpenAPI 定義（JSON）から JMX を生成
./jmeter_api_tool.sh generate -i openapi.json -c input/openapi_config.csv   # 接続先・ヘッダ・トークンを OpenAPI 用設定ファイルで指定
./jmeter_api_tool.sh run -t output/api_test.jmx             # CLI 実行（result/JMX名_日時/）
./jmeter_api_tool.sh run -t output/api_test.jmx --threads 10 --loops 100   # 大量実行
./jmeter_api_tool.sh manual                                 # 利用方法を出力（docs/）
```

### Windows 11

```powershell
.\jmeter_api_tool.bat init
.\jmeter_api_tool.bat generate -i input\api_definition.xlsx
.\jmeter_api_tool.bat generate -i input\api_definition.xlsx -JmxName order_test.jmx
.\jmeter_api_tool.bat generate -i openapi.json -Seed 12345
.\jmeter_api_tool.bat generate -i openapi.json -c input\openapi_config.csv
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

- 設定は Excel / CSV と同じく `-s` または「入力ファイル名_settings.csv」から読み込みます。設定ファイルで空欄の `base_url`（`servers[0].url`。無い場合は `http://localhost:8080`）、`test_plan_name`、`output_file`（`<タイトル>.jmx`）は OpenAPI 定義の値を使います（`output_file` は `-n` / `-JmxName` を指定するとその値が優先）。
- 認証ヘッダ（Authorization・API-Key）はランダム値では通らないため生成しません。設定 `common_headers`、または下記の OpenAPI 用設定ファイル（`HEADERS` / `TOKEN_DIR`）で指定してください。
- OpenAPI 用のオプション（sh / PowerShell）: `-c` / `-OpenApiConfig`（OpenAPI 用設定ファイル。下記）、`--seed` / `-Seed`（乱数シード）、`--any-method` / `-AnyMethod`、`--required-only` / `-RequiredOnly`、`--use-examples` / `-UseExamples`、`--placeholder-value` / `-PlaceholderValue`（意味は JMeter_Tools と同じ）。
- GET / DELETE のリクエストボディは Excel / CSV と同じ規則で送信しません（警告を表示）。YAML・Swagger 2.0・外部ファイルへの `$ref` は対象外です。

### OpenAPI 用設定ファイル（`-c` / `-OpenApiConfig`）

OpenAPI 定義から生成する場合だけ使う設定ファイルです。`init` で作成される `input/openapi_config.csv`（設定 CSV と同じ列: 設定キー / 設定項目 / 設定値 / 既定値 / 説明）の「設定値」を記入し、`generate -i openapi.json -c input/openapi_config.csv` のように指定します。
設定値が空欄の項目は使われず、テンプレートのままなら JMX は指定しない場合と同じです。記入した項目は、設定ファイル（`-s`）・OpenAPI 定義の値より優先します。

| 設定キー | 内容 |
|---|---|
| `PROTOCOL` / `HOST` / `PORT` | 共通変数（User Defined Variables）PROTOCOL / HOST / PORT の既定値。記入した項目だけ `base_url`（`servers[0].url`）の接続先を置き換え、パス（例: `/v1`）はそのまま使います。PORT が空欄の場合はポート番号を付けないため（未設定。http は 80 / https は 443 で接続）、ホスト名だけでも指定できます（3 項目とも空欄なら `base_url` のまま）。実行時の `-JPROTOCOL=` `-JHOST=` `-JPORT=` も従来どおり使えます |
| `HEADERS` | 全 API に付ける HTTP ヘッダ（「名前: 値」を改行区切り。例: `X-Api-Key: キーの値`）。共通ヘッダ（HTTP Header Manager）に追加し、設定 `common_headers` と同じ名前のヘッダはこちらを使います |
| `TOKEN_DIR` | アクセストークンのフォルダ。指定すると共通ヘッダに `Authorization: Bearer ${ACCESS_TOKEN}` を追加し、JMeter の実行中に API を呼び出すたびに、このフォルダ直下の `*.token` のうち **更新日時が最も新しいファイル** の内容（前後の空白・改行・BOM を除く）を `${ACCESS_TOKEN}` に設定します（`HEADERS` に Authorization を書いた場合はその値を使用） |

- `TOKEN_DIR` の書き方（Windows / Linux 共通）: 区切りは `\` と `/` のどちらでも可。相対パスは JMX のフォルダ基準（例: `tokens`、`..\tokens`）。Windows と Linux のフォルダを `;` で区切って書くと、実行する環境に存在する方を使うため、同じ JMX を両方で使えます（例: `C:\work\tokens;/home/user/tokens`）。実行時に `-JACCESS_TOKEN_DIR=フォルダ` で変更することもできます。
- トークンは JMX には書き込まず、JMeter の実行時に読み込みます。新しい `.token` ファイルを置く・ファイルを書き換えると、実行中でも次の API 呼び出しから切り替わります（フォルダの確認は 1 秒に 1 回）。拡張子が `.token` 以外のファイルは使いません。
- 生成時には、そのとき最新の `.token` ファイル（またはフォルダ・ファイルが無いことの警告）を表示します。

## 生成される JMX の主な内容

- 既定は 1 スレッド × 1 回。`-JTHREADS` `-JLOOPS` `-JRAMP_UP`（`-JDURATION`）で大量件数の負荷試験に切り替え
- 接続先は `${__P(HOST,...)}` 形式の共通変数（`-JHOST=` で環境切替）
- API ごと: ヘッダ、定数タイマ（待機）、前処理・後処理（JSR223 / Groovy）、正規表現抽出、JSON 抽出、ステータス・本文・応答時間の検証
- 共通: 連番（Counter）、乱数、CSV データ読込、Excel データ読込（Apache POI 不要）、クッキー管理、共通前処理・後処理、スループット制御（既定は無効）
- アクセストークン: 設定 `access_token_dir` を指定すると、そのフォルダ配下のファイル（`access_token_file`、既定 `access_token.txt`）からトークンを読み込み、共通ヘッダに `Authorization: Bearer ${ACCESS_TOKEN}` を付けます（実行中のファイル更新に追従。`-JACCESS_TOKEN_DIR` / `-JACCESS_TOKEN_FILE` で変更可）。`access_token_file` を `*.token` のように `*` を含めて書くと最も新しいファイルを使い、フォルダは `\` / `/` どちらの区切りでも、`;` 区切りで複数（存在する最初のフォルダを使用）でも書けます
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
