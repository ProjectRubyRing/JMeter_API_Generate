<#
.SYNOPSIS
  jmeter_api_tool.ps1 v1.0.0 （Windows 11 用 / Windows PowerShell 5.1・PowerShell 7）

.DESCRIPTION
  REST API の定義（Excel / CSV / OpenAPI）から Apache JMeter 5.6.3 用の JMX を生成し、
  CLI モードで実行して JTL・HTML レポート・ログを出力するツールです。
  追加モジュール・Excel のインストールは不要です（xlsx は ZIP + XML として直接読み書きします）。
  ※ 処理内容は RHEL 版 jmeter_api_tool.sh と同一です。

  実行ポリシーで .ps1 の実行が禁止されている場合は、同じフォルダの jmeter_api_tool.bat から実行してください。

.EXAMPLE
  .\jmeter_api_tool.bat init
  .\jmeter_api_tool.bat generate -i input\api_definition.xlsx
  .\jmeter_api_tool.bat run -t output\api_test.jmx -Threads 10 -Loops 100
#>
param(
    [string]$Command = 'help',
    [Alias('i')][string]$InputFile = '',
    [Alias('s')][string]$SettingsFile = '',
    [Alias('o')][string]$OutDir = '',
    [switch]$Split,
    [Alias('f')][switch]$Force,
    [Alias('t')][string]$Jmx = '',
    [string]$Threads = '',
    [string]$RampUp = '',
    [string]$Loops = '',
    [string]$Duration = '',
    [string[]]$Prop = @(),
    [string]$Heap = '',
    [string]$JMeterHome = '',
    [switch]$FailOnError,
    [string]$ExtraArgs = '',
    [string]$Seed = '',
    [string]$AnyMethod = 'GET',
    [switch]$RequiredOnly,
    [switch]$UseExamples,
    [string]$PlaceholderValue = '3000',
    [Alias('n')][string]$JmxName = '',
    [Alias('c')][string]$OpenApiConfig = ''
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:ToolVersion = '1.0.0'
$script:Warnings = New-Object System.Collections.Generic.List[string]
try { Add-Type -AssemblyName System.IO.Compression } catch { }
try { Add-Type -AssemblyName System.IO.Compression.FileSystem } catch { }
try { [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance) } catch { }

$script:SpecJson = @'
{
 "tool": {
  "name": "jmeter_api_tool",
  "version": "1.0.0",
  "jmeter": "5.6.3"
 },
 "settings": [
  {
   "category": "基本",
   "key": "test_plan_name",
   "label": "テストプラン名",
   "default": "APIテスト",
   "type": "text",
   "desc": "JMeter 上に表示されるテストプラン名"
  },
  {
   "category": "基本",
   "key": "output_file",
   "label": "出力JMXファイル名",
   "default": "api_test.jmx",
   "type": "text",
   "desc": "生成する JMX のファイル名（英数字・_・-・. 以外の文字は _ に置き換えます）"
  },
  {
   "category": "基本",
   "key": "base_url",
   "label": "ベースURL",
   "default": "",
   "type": "text",
   "desc": "API一覧の URL が「/」で始まる場合に先頭に付ける URL（例: https://api.example.com/v1）"
  },
  {
   "category": "基本",
   "key": "content_encoding",
   "label": "文字コード",
   "default": "UTF-8",
   "type": "text",
   "desc": "リクエストの文字コード"
  },
  {
   "category": "負荷",
   "key": "threads",
   "label": "スレッド数",
   "default": "1",
   "type": "int",
   "desc": "同時に実行する仮想ユーザー数。-JTHREADS"
  },
  {
   "category": "負荷",
   "key": "ramp_up",
   "label": "ランプアップ(秒)",
   "default": "1",
   "type": "int",
   "desc": "全スレッドが起動し終わるまでの秒数。-JRAMP_UP"
  },
  {
   "category": "負荷",
   "key": "loops",
   "label": "ループ回数",
   "default": "1",
   "type": "int",
   "desc": "各スレッドが API 一覧を繰り返す回数。-1 = 無限（実行時間と併用）。-JLOOPS"
  },
  {
   "category": "負荷",
   "key": "duration",
   "label": "実行時間(秒)",
   "default": "",
   "type": "int",
   "desc": "指定するとスケジューラを有効にし、この秒数で終了（ループ回数は -1 を推奨）。-JDURATION"
  },
  {
   "category": "負荷",
   "key": "on_sample_error",
   "label": "エラー時の動作",
   "default": "continue",
   "type": "enum",
   "desc": "continue = 続行 / startnextloop = 次のループへ / stopthread = スレッド停止 / stoptest = テスト停止 / stoptestnow = 即時停止",
   "choices": [
    "continue",
    "startnextloop",
    "stopthread",
    "stoptest",
    "stoptestnow"
   ]
  },
  {
   "category": "負荷",
   "key": "throughput_timer",
   "label": "スループット制御",
   "default": "DISABLE",
   "type": "onoff",
   "desc": "1 分あたりのリクエスト数を一定に保つ Constant Throughput Timer",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "負荷",
   "key": "throughput_per_min",
   "label": "目標スループット(件/分)",
   "default": "600",
   "type": "int",
   "desc": "スループット制御の目標値（メイン処理全体の合計）。-JTHROUGHPUT"
  },
  {
   "category": "HTTP",
   "key": "http_defaults",
   "label": "HTTPリクエスト初期値",
   "default": "ON",
   "type": "onoff",
   "desc": "文字コード・タイムアウトの共通設定（HTTP Request Defaults）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "HTTP",
   "key": "connect_timeout",
   "label": "接続タイムアウト(ms)",
   "default": "10000",
   "type": "int",
   "desc": "-JCONNECT_TIMEOUT"
  },
  {
   "category": "HTTP",
   "key": "response_timeout",
   "label": "応答タイムアウト(ms)",
   "default": "30000",
   "type": "int",
   "desc": "-JRESPONSE_TIMEOUT"
  },
  {
   "category": "HTTP",
   "key": "common_headers",
   "label": "共通ヘッダ",
   "default": "",
   "type": "lines",
   "desc": "全 API に付けるヘッダ。「名前: 値」を改行区切り",
   "template": "Accept: application/json\nUser-Agent: JMeter-API-Test/1.0\nX-Client-Id: ${CLIENT_ID}"
  },
  {
   "category": "HTTP",
   "key": "access_token_dir",
   "label": "アクセストークンのフォルダ",
   "default": "",
   "type": "text",
   "desc": "指定すると、このフォルダ配下のファイルからアクセストークンを読み込み、共通ヘッダに Authorization: Bearer ${ACCESS_TOKEN} を付けます（共通ヘッダに Authorization がある場合はその値を使用）。空欄 = 使わない。JMX からの相対パスまたは絶対パス（区切りは \\ と / のどちらでも可。; で区切って複数書くと、存在する最初のフォルダを使用）。-JACCESS_TOKEN_DIR"
  },
  {
   "category": "HTTP",
   "key": "access_token_file",
   "label": "アクセストークンのファイル名",
   "default": "access_token.txt",
   "type": "text",
   "desc": "アクセストークンを書いたファイルの名前（フォルダからの相対パス）。*.token のように * を含めると、一致するファイルのうち更新日時が最も新しいファイルを使います。前後の空白・改行は除き、実行中にファイルが更新されると読み直します。-JACCESS_TOKEN_FILE"
  },
  {
   "category": "HTTP",
   "key": "follow_redirects",
   "label": "リダイレクト追従",
   "default": "true",
   "type": "bool",
   "desc": "true = 3xx 応答の Location を自動で追従",
   "choices": [
    "true",
    "false"
   ]
  },
  {
   "category": "HTTP",
   "key": "keep_alive",
   "label": "KeepAlive",
   "default": "true",
   "type": "bool",
   "desc": "true = 接続を再利用する",
   "choices": [
    "true",
    "false"
   ]
  },
  {
   "category": "HTTP",
   "key": "default_wait_ms",
   "label": "既定の待機時間(ms)",
   "default": "0",
   "type": "int",
   "desc": "API一覧の「待機時間」が空欄の API に使う値（0 = 待機なし）"
  },
  {
   "category": "クッキー",
   "key": "cookie_manager",
   "label": "クッキー管理",
   "default": "ON",
   "type": "onoff",
   "desc": "HTTP Cookie Manager（Set-Cookie を保存して以降のリクエストで送信）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "クッキー",
   "key": "cookie_clear_each_iteration",
   "label": "反復ごとにクリア",
   "default": "true",
   "type": "bool",
   "desc": "true = ループのたびにクッキーを消去（毎回新しいセッション）",
   "choices": [
    "true",
    "false"
   ]
  },
  {
   "category": "クッキー",
   "key": "cookies",
   "label": "事前設定クッキー",
   "default": "",
   "type": "lines",
   "desc": "最初から送信するクッキー。「名前=値」を改行区切り",
   "template": "LANG=ja"
  },
  {
   "category": "クッキー",
   "key": "cookie_domain",
   "label": "クッキーのドメイン",
   "default": "",
   "type": "text",
   "desc": "事前設定クッキーのドメイン。空欄 = 1 つ目の接続先（${HOST}）"
  },
  {
   "category": "変数",
   "key": "user_variables",
   "label": "共通変数",
   "default": "",
   "type": "lines",
   "desc": "「名前=値」を改行区切り。${名前} で参照でき、実行時に -J名前=値 で変更可",
   "template": "ENV=staging\nCLIENT_ID=jmeter-client"
  },
  {
   "category": "変数",
   "key": "counter",
   "label": "連番",
   "default": "ON",
   "type": "onoff",
   "desc": "反復ごとに連番を設定する Counter",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "変数",
   "key": "counter_name",
   "label": "連番の変数名",
   "default": "SEQ",
   "type": "text",
   "desc": "連番を入れる変数名"
  },
  {
   "category": "変数",
   "key": "counter_start",
   "label": "連番の開始値",
   "default": "1",
   "type": "int",
   "desc": ""
  },
  {
   "category": "変数",
   "key": "counter_increment",
   "label": "連番の増分",
   "default": "1",
   "type": "int",
   "desc": ""
  },
  {
   "category": "変数",
   "key": "counter_format",
   "label": "連番の書式",
   "default": "000000",
   "type": "text",
   "desc": "0 埋めの桁数など（Java の DecimalFormat 形式）。空欄 = 書式なし"
  },
  {
   "category": "変数",
   "key": "counter_per_thread",
   "label": "スレッドごとの連番",
   "default": "false",
   "type": "bool",
   "desc": "true = スレッドごとに開始値から採番 / false = 全スレッドで一意",
   "choices": [
    "true",
    "false"
   ]
  },
  {
   "category": "変数",
   "key": "random_variable",
   "label": "乱数",
   "default": "ON",
   "type": "onoff",
   "desc": "反復ごとに乱数を設定する Random Variable",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "変数",
   "key": "random_name",
   "label": "乱数の変数名",
   "default": "RANDOM_NO",
   "type": "text",
   "desc": ""
  },
  {
   "category": "変数",
   "key": "random_min",
   "label": "乱数の最小値",
   "default": "1",
   "type": "int",
   "desc": ""
  },
  {
   "category": "変数",
   "key": "random_max",
   "label": "乱数の最大値",
   "default": "99999",
   "type": "int",
   "desc": ""
  },
  {
   "category": "CSVデータ",
   "key": "csv_data",
   "label": "CSVデータ読込",
   "default": "ON",
   "type": "onoff",
   "desc": "CSV Data Set Config（1 反復につき 1 行を読み込み、列を変数に設定）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "CSVデータ",
   "key": "csv_file",
   "label": "CSVファイル",
   "default": "data/testdata.csv",
   "type": "text",
   "desc": "JMX からの相対パスまたは絶対パス。-JCSV_FILE"
  },
  {
   "category": "CSVデータ",
   "key": "csv_variables",
   "label": "CSVの変数名",
   "default": "",
   "type": "text",
   "desc": "カンマ区切り。空欄 = 生成時に CSV ファイルの 1 行目を読み取って設定"
  },
  {
   "category": "CSVデータ",
   "key": "csv_encoding",
   "label": "CSVの文字コード",
   "default": "UTF-8",
   "type": "text",
   "desc": ""
  },
  {
   "category": "CSVデータ",
   "key": "csv_delimiter",
   "label": "CSVの区切り文字",
   "default": ",",
   "type": "text",
   "desc": ""
  },
  {
   "category": "CSVデータ",
   "key": "csv_recycle",
   "label": "終端で先頭に戻る",
   "default": "true",
   "type": "bool",
   "desc": "true = 最終行の次は先頭行に戻る",
   "choices": [
    "true",
    "false"
   ]
  },
  {
   "category": "CSVデータ",
   "key": "csv_stop_thread",
   "label": "終端でスレッド停止",
   "default": "false",
   "type": "bool",
   "desc": "true = 最終行まで読んだらスレッドを停止（csv_recycle=false と併用）",
   "choices": [
    "true",
    "false"
   ]
  },
  {
   "category": "CSVデータ",
   "key": "csv_share_mode",
   "label": "CSVの共有範囲",
   "default": "all",
   "type": "enum",
   "desc": "all = 全スレッドで共有 / group = スレッドグループ単位 / thread = スレッドごとに先頭から",
   "choices": [
    "all",
    "group",
    "thread"
   ]
  },
  {
   "category": "Excelデータ",
   "key": "excel_data",
   "label": "Excelデータ読込",
   "default": "ON",
   "type": "onoff",
   "desc": "Excel（.xlsx）の読み込み。Apache POI などの追加ライブラリは不要",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "Excelデータ",
   "key": "excel_file",
   "label": "Excelファイル",
   "default": "data/testdata.xlsx",
   "type": "text",
   "desc": "JMX からの相対パスまたは絶対パス。-JEXCEL_FILE"
  },
  {
   "category": "Excelデータ",
   "key": "excel_sheet",
   "label": "シート名",
   "default": "",
   "type": "text",
   "desc": "空欄 = 先頭のシート。-JEXCEL_SHEET"
  },
  {
   "category": "前後処理",
   "key": "setup_teardown",
   "label": "開始・終了処理",
   "default": "ON",
   "type": "onoff",
   "desc": "テスト開始・終了時のログ出力と後片付け（setUp / tearDown Thread Group）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "前後処理",
   "key": "common_preprocessor",
   "label": "共通前処理",
   "default": "ON",
   "type": "onoff",
   "desc": "全 API の直前に実行する JSR223 PreProcessor（REQUEST_ID・TIMESTAMP 等を生成）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "前後処理",
   "key": "common_postprocessor",
   "label": "共通後処理",
   "default": "ON",
   "type": "onoff",
   "desc": "全 API の直後に実行する JSR223 PostProcessor（LAST_STATUS 等を設定）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "結果出力",
   "key": "result_dir",
   "label": "結果出力フォルダ",
   "default": "~/result",
   "type": "text",
   "desc": "JTL・結果ログの出力先。~/ = JMX のフォルダ。-JRESULT_DIR（run コマンドは自動設定）"
  },
  {
   "category": "結果出力",
   "key": "view_results_tree",
   "label": "結果ツリー",
   "default": "ON",
   "type": "onoff",
   "desc": "View Results Tree（GUI での確認用）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "結果出力",
   "key": "summary_report",
   "label": "集計レポート",
   "default": "ON",
   "type": "onoff",
   "desc": "Summary Report（GUI での確認用）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "結果出力",
   "key": "aggregate_report",
   "label": "統計レポート",
   "default": "ON",
   "type": "onoff",
   "desc": "Aggregate Report（GUI での確認用。パーセンタイル）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "結果出力",
   "key": "jtl_writer",
   "label": "JTL出力(全件)",
   "default": "ON",
   "type": "onoff",
   "desc": "全リクエストの結果を result_all.jtl（CSV）に出力",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "結果出力",
   "key": "error_writer",
   "label": "エラー詳細出力",
   "default": "ON",
   "type": "onoff",
   "desc": "失敗したリクエストのみ error_detail.xml（リクエスト・レスポンス全文）に出力",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "結果出力",
   "key": "result_log",
   "label": "結果ログ出力",
   "default": "ON",
   "type": "onoff",
   "desc": "1 リクエスト 1 行の結果ログ api_result.log を出力（JSR223 Listener）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "結果出力",
   "key": "log_response",
   "label": "結果ログの本文出力",
   "default": "error",
   "type": "enum",
   "desc": "none = 出力しない / error = 失敗時のみ / all = 全件。-JLOG_RESPONSE",
   "choices": [
    "none",
    "error",
    "all"
   ]
  },
  {
   "category": "結果出力",
   "key": "log_max_chars",
   "label": "本文の最大文字数",
   "default": "1000",
   "type": "int",
   "desc": "結果ログに出力する本文の最大文字数。-JLOG_MAX_CHARS"
  },
  {
   "category": "結果出力",
   "key": "debug_sampler",
   "label": "変数確認(デバッグ)",
   "default": "DISABLE",
   "type": "onoff",
   "desc": "Debug Sampler（全変数を結果ツリーに表示）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  }
 ],
 "openapi_config": [
  {
   "category": "接続先",
   "key": "PROTOCOL",
   "label": "プロトコル",
   "default": "",
   "desc": "共通変数 PROTOCOL の既定値（http / https）。空欄 = OpenAPI 定義（servers[0].url）または設定 base_url の値。実行時は -JPROTOCOL=値 で変更できます"
  },
  {
   "category": "接続先",
   "key": "HOST",
   "label": "ホスト名",
   "default": "",
   "desc": "共通変数 HOST の既定値（例: api.example.com）。空欄 = OpenAPI 定義（servers[0].url）または設定 base_url の値。実行時は -JHOST=値 で変更できます"
  },
  {
   "category": "接続先",
   "key": "PORT",
   "label": "ポート番号",
   "default": "",
   "desc": "共通変数 PORT の既定値（1～65535）。空欄 = 未設定（ポート番号なし。http は 80 / https は 443 で接続）。PROTOCOL と HOST も空欄の場合は OpenAPI 定義（servers[0].url）または設定 base_url の値。実行時は -JPORT=値 で変更できます"
  },
  {
   "category": "HTTPヘッダ",
   "key": "HEADERS",
   "label": "HTTPヘッダ",
   "default": "",
   "desc": "全 API に付ける HTTP ヘッダ（共通ヘッダの HTTP Header Manager に追加）。「名前: 値」を改行区切り（例: X-Api-Key: キーの値）。設定 common_headers と同じ名前のヘッダはこちらを使います"
  },
  {
   "category": "HTTPヘッダ",
   "key": "API_HEADERS",
   "label": "各APIのHTTPヘッダ（初期値）",
   "default": "Accept: application/json\nTransaction-ID: XXXXXXXXX\nTransaction-Class: 1\nUser-ID: xxxxxxxx\nAPI-Key: XXXXXXX",
   "desc": "OpenAPI 定義から生成する各 API のヘッダ（API ごとの HTTP Header Manager）の初期値（生成後に JMeter で API ごとに変更できます）。「名前: 値」を改行区切り。OpenAPI 定義から求めた同じ名前のヘッダ（Accept など）はこちらの値を使います。空欄 = 既定値。付けない場合は「なし」と記入"
  },
  {
   "category": "アクセストークン",
   "key": "TOKEN_DIR",
   "label": "アクセストークンのフォルダ",
   "default": "",
   "desc": "指定すると、共通ヘッダに Authorization: Bearer ${ACCESS_TOKEN} を追加し、JMeter の実行中に API を呼び出すたびに、このフォルダ直下の拡張子 .token のファイルのうち更新日時が最も新しいファイルの内容を ${ACCESS_TOKEN} に設定します（HEADERS に Authorization を書いた場合はその値を使用）。区切りは \\ と / のどちらでも可、相対パスは JMX のフォルダ基準。Windows と Linux のフォルダを ; で区切って書くと、存在する方を使います（例: C:\\work\\tokens;/home/user/tokens）。実行時は -JACCESS_TOKEN_DIR=フォルダ で変更できます"
  },
  {
   "category": "OpenAPI定義の変数",
   "key": "PUBLIC_BASE_URL",
   "label": "ベースURL（${public_base_url}）",
   "default": "",
   "desc": "OpenAPI 定義（servers[0].url）のプレースホルダ ${public_base_url} に入れる値。URL（例: https://api.example.com/v1）は接続先とパスの先頭に、パス（例: /v1）はパスの先頭（コンテキストルート）に使います。空欄 = 未設定（${public_base_url} を取り除き、パスに付けません）"
  },
  {
   "category": "OpenAPI定義の変数",
   "key": "NAME_PREFIX",
   "label": "名前の接頭辞（${name_prefix}）",
   "default": "",
   "desc": "OpenAPI 定義（info.title）のプレースホルダ ${name_prefix} に入れる値（例: dev）。テストプラン名と出力する JMX のファイル名に使われます。空欄 = 未設定（${name_prefix} を隣の区切り文字（- _ . / 空白）ごと取り除きます。例: ${name_prefix}-api → api）"
  }
 ],
 "api_columns": [
  {
   "key": "no",
   "header": "No",
   "aliases": [
    "番号",
    "no."
   ],
   "required": false,
   "width": 6,
   "hint": "#説明",
   "desc": "実行順の番号（ラベル APIxx に使用）。空欄は自動採番。# で始まる行は読み飛ばす",
   "example": "1"
  },
  {
   "key": "enabled",
   "header": "有効",
   "aliases": [
    "enabled",
    "有効/無効",
    "有効フラグ"
   ],
   "required": false,
   "width": 6,
   "hint": "Y / N",
   "desc": "Y = 有効、N = 無効化した状態で生成。空欄 = Y",
   "example": "Y"
  },
  {
   "key": "name",
   "header": "API名",
   "aliases": [
    "name",
    "名前",
    "api"
   ],
   "required": false,
   "width": 22,
   "hint": "任意",
   "desc": "HTTP リクエストの名前（レポートのラベル）。空欄はメソッドとパス",
   "example": "ログイン"
  },
  {
   "key": "method",
   "header": "メソッド",
   "aliases": [
    "method",
    "httpメソッド",
    "httpmethod"
   ],
   "required": true,
   "width": 10,
   "hint": "GET / POST",
   "desc": "HTTP メソッド（GET / POST。PUT / PATCH / DELETE も可）",
   "example": "POST"
  },
  {
   "key": "url",
   "header": "URL",
   "aliases": [
    "url",
    "エンドポイント",
    "endpoint"
   ],
   "required": true,
   "width": 42,
   "hint": "必須（https://～）",
   "desc": "完全な URL、またはベースURL 設定時は「/」で始まるパス。${変数} 可。? 以降のクエリも可",
   "example": "https://api.example.com/v1/users"
  },
  {
   "key": "query",
   "header": "クエリパラメータ",
   "aliases": [
    "query",
    "params",
    "クエリ",
    "パラメータ"
   ],
   "required": false,
   "width": 28,
   "hint": "名前=値（改行 / & 区切り）",
   "desc": "URL のクエリパラメータ",
   "example": "page=1\nsize=20"
  },
  {
   "key": "headers",
   "header": "ヘッダ",
   "aliases": [
    "headers",
    "header",
    "リクエストヘッダ",
    "ヘッダ情報"
   ],
   "required": false,
   "width": 34,
   "hint": "名前: 値（改行区切り）",
   "desc": "この API だけに付けるヘッダ（全 API 共通のヘッダは設定 common_headers）",
   "example": "Authorization: Bearer ${AUTH_TOKEN}"
  },
  {
   "key": "body_type",
   "header": "ボディ形式",
   "aliases": [
    "body_type",
    "bodytype",
    "形式"
   ],
   "required": false,
   "width": 11,
   "hint": "auto / json / form / raw",
   "desc": "リクエストボディの形式。空欄 = auto（自動判定）",
   "example": "json"
  },
  {
   "key": "body",
   "header": "リクエストボディ",
   "aliases": [
    "body",
    "requestbody",
    "ボディ",
    "postパラメータ"
   ],
   "required": false,
   "width": 40,
   "hint": "JSON 等 / 名前=値",
   "desc": "POST 等で送る本文。form 形式は「名前=値」を改行 / & 区切り",
   "example": "{\"userId\":\"${USER_ID}\"}"
  },
  {
   "key": "wait_ms",
   "header": "待機時間(ms)",
   "aliases": [
    "wait",
    "wait_ms",
    "待機時間",
    "待機"
   ],
   "required": false,
   "width": 11,
   "hint": "数値（ms）",
   "desc": "API 呼び出し前に待つ時間（Constant Timer）。空欄 = 設定 default_wait_ms、0 = 待機なし",
   "example": "1000"
  },
  {
   "key": "expected_status",
   "header": "期待ステータス",
   "aliases": [
    "status",
    "expected_status",
    "ステータス"
   ],
   "required": false,
   "width": 13,
   "hint": "200 / 200,201",
   "desc": "期待する HTTP ステータス（カンマ区切りで複数可）。4xx / 5xx を書くと異常系テストとして扱う",
   "example": "200"
  },
  {
   "key": "expected_text",
   "header": "応答検証文字列",
   "aliases": [
    "expected_text",
    "assert",
    "検証文字列"
   ],
   "required": false,
   "width": 22,
   "hint": "文字列（改行区切り）",
   "desc": "応答本文に含まれるべき文字列。複数書くとすべて含む場合に成功",
   "example": "\"result\":\"OK\""
  },
  {
   "key": "max_time_ms",
   "header": "応答時間上限(ms)",
   "aliases": [
    "max_time",
    "max_time_ms",
    "応答時間上限"
   ],
   "required": false,
   "width": 13,
   "hint": "数値（ms）",
   "desc": "この時間を超えたら失敗（Duration Assertion）",
   "example": "3000"
  },
  {
   "key": "regex_extract",
   "header": "正規表現抽出",
   "aliases": [
    "regex",
    "regex_extract",
    "正規表現"
   ],
   "required": false,
   "width": 34,
   "hint": "変数名=正規表現（改行区切り）",
   "desc": "応答本文から ( ) 内を抽出して変数に設定。見つからない場合は NOT_FOUND",
   "example": "TOKEN=\"token\"\\s*:\\s*\"([^\"]+)\""
  },
  {
   "key": "json_extract",
   "header": "JSON抽出",
   "aliases": [
    "json",
    "json_extract",
    "jsonpath"
   ],
   "required": false,
   "width": 28,
   "hint": "変数名=JSONPath（改行区切り）",
   "desc": "応答 JSON から値を抽出して変数に設定。見つからない場合は NOT_FOUND",
   "example": "USER_ID=$.data.id"
  },
  {
   "key": "pre_script",
   "header": "前処理スクリプト",
   "aliases": [
    "pre",
    "pre_script",
    "前処理"
   ],
   "required": false,
   "width": 36,
   "hint": "Groovy",
   "desc": "API 呼び出しの直前に実行する Groovy（JSR223 PreProcessor）",
   "example": "vars.put('X', '1')"
  },
  {
   "key": "post_script",
   "header": "後処理スクリプト",
   "aliases": [
    "post",
    "post_script",
    "後処理"
   ],
   "required": false,
   "width": 36,
   "hint": "Groovy",
   "desc": "API 呼び出しの直後に実行する Groovy（JSR223 PostProcessor）。抽出の後に実行",
   "example": "log.info(prev.getResponseCode())"
  },
  {
   "key": "comment",
   "header": "備考",
   "aliases": [
    "comment",
    "コメント",
    "メモ"
   ],
   "required": false,
   "width": 24,
   "hint": "任意",
   "desc": "JMeter のコメント欄に設定",
   "example": ""
  }
 ],
 "sample_apis": [
  {
   "no": "1",
   "enabled": "Y",
   "name": "ログイン(POST JSON)",
   "method": "POST",
   "url": "https://postman-echo.com/post",
   "query": "",
   "headers": "Content-Type: application/json\nX-Correlation-ID: ${REQUEST_ID}\nX-Signature: ${SIGNATURE}",
   "body_type": "json",
   "body": "{\n  \"userId\": \"${USER_ID}\",\n  \"password\": \"${PASSWORD}\",\n  \"requestNo\": \"${SEQ}\",\n  \"timestamp\": \"${TIMESTAMP}\"\n}",
   "wait_ms": "500",
   "expected_status": "200",
   "expected_text": "${USER_ID}",
   "max_time_ms": "5000",
   "regex_extract": "ECHO_URL=\"url\"\\s*:\\s*\"([^\"]+)\"",
   "json_extract": "AUTH_TOKEN=$.json.userId\nLOGIN_SEQ=$.json.requestNo",
   "pre_script": "// 署名の生成例: USER_ID + TIMESTAMP の SHA-256（16進）を変数 SIGNATURE に設定\nString src = (vars.get('USER_ID') ?: '') + (vars.get('TIMESTAMP') ?: '')\ndef digest = java.security.MessageDigest.getInstance('SHA-256').digest(src.getBytes('UTF-8'))\nvars.put('SIGNATURE', digest.encodeHex().toString())",
   "post_script": "// 応答 JSON の解析例: 値を取り出して次の API で使う変数に設定\ntry {\n    def json = new groovy.json.JsonSlurper().parseText(prev.getResponseDataAsString())\n    vars.put('LOGIN_USER', String.valueOf(json?.json?.userId))\n} catch (Exception e) {\n    log.warn('ログイン応答の解析に失敗しました: ' + e.getMessage())\n}",
   "comment": "JSON を POST する例（postman-echo.com は送信内容をそのまま返す）"
  },
  {
   "no": "2",
   "enabled": "Y",
   "name": "商品検索(GET)",
   "method": "GET",
   "url": "https://postman-echo.com/get",
   "query": "keyword=${KEYWORD}\nitemCode=${ITEM_CODE}\nrand=${RANDOM_NO}\npage=1",
   "headers": "Authorization: Bearer ${AUTH_TOKEN}",
   "body_type": "",
   "body": "",
   "wait_ms": "1000",
   "expected_status": "200",
   "expected_text": "${ITEM_CODE}",
   "max_time_ms": "5000",
   "regex_extract": "",
   "json_extract": "SEARCH_KEYWORD=$.args.keyword",
   "pre_script": "",
   "post_script": "",
   "comment": "GET とクエリパラメータの例（CSV・Excel・乱数の変数を使用）"
  },
  {
   "no": "3",
   "enabled": "Y",
   "name": "注文登録(POST form)",
   "method": "POST",
   "url": "https://postman-echo.com/post?source=jmeter",
   "query": "channel=web",
   "headers": "Authorization: Bearer ${AUTH_TOKEN}",
   "body_type": "form",
   "body": "itemCode=${ITEM_CODE}\nitemName=${ITEM_NAME}\nquantity=${QUANTITY}\norderNo=ORD${SEQ}",
   "wait_ms": "1000",
   "expected_status": "200,201",
   "expected_text": "ORD${SEQ}",
   "max_time_ms": "5000",
   "regex_extract": "ORDER_NO=\"orderNo\"\\s*:\\s*\"([^\"]+)\"",
   "json_extract": "",
   "pre_script": "",
   "post_script": "// 抽出した値の利用例（jmeter.log に出力）\nlog.info('注文番号=' + vars.get('ORDER_NO') + ' / ステータス=' + prev.getResponseCode())",
   "comment": "form 形式の POST の例（POST のクエリパラメータは URL に付与）"
  },
  {
   "no": "4",
   "enabled": "Y",
   "name": "クッキー確認(GET)",
   "method": "GET",
   "url": "https://postman-echo.com/cookies/set",
   "query": "session_id=S${SEQ}",
   "headers": "",
   "body_type": "",
   "body": "",
   "wait_ms": "500",
   "expected_status": "200",
   "expected_text": "session_id",
   "max_time_ms": "5000",
   "regex_extract": "",
   "json_extract": "SESSION_COOKIE=$.cookies.session_id",
   "pre_script": "",
   "post_script": "",
   "comment": "Set-Cookie とリダイレクト追従の例（クッキー管理を使用）"
  }
 ],
 "sample_csv": {
  "header": [
   "USER_ID",
   "PASSWORD",
   "KEYWORD"
  ],
  "rows": [
   [
    "user001",
    "Passw0rd!01",
    "テスト"
   ],
   [
    "user002",
    "Passw0rd!02",
    "JMeter"
   ],
   [
    "user003",
    "Passw0rd!03",
    "負荷試験"
   ],
   [
    "user004",
    "Passw0rd!04",
    "API"
   ],
   [
    "user005",
    "Passw0rd!05",
    "サンプル"
   ]
  ]
 },
 "sample_excel": {
  "sheet": "data",
  "header": [
   "ITEM_CODE",
   "ITEM_NAME",
   "QUANTITY"
  ],
  "rows": [
   [
    "A001",
    "りんご",
    "1"
   ],
   [
    "A002",
    "みかん",
    "2"
   ],
   [
    "B001",
    "ぶどう",
    "3"
   ],
   [
    "B002",
    "もも",
    "1"
   ],
   [
    "C001",
    "メロン",
    "5"
   ]
  ]
 },
 "names": {
  "udv": "[必須] 共通変数 (User Defined Variables)",
  "http_defaults": "[任意] HTTPリクエスト初期値 (HTTP Request Defaults)",
  "common_headers": "[任意] 共通ヘッダ (HTTP Header Manager)",
  "cookie": "[任意] クッキー管理 (HTTP Cookie Manager)",
  "setup_tg": "[任意] 事前処理 (setUp Thread Group)",
  "setup_start": "[任意] テスト開始処理 (JSR223 Sampler)",
  "excel_loader": "[任意] Excelデータ読込 (JSR223 Sampler)",
  "main_tg": "[必須] メイン処理 (Thread Group)",
  "counter": "[任意] 連番 @@NAME@@ (Counter)",
  "random": "[任意] 乱数 @@NAME@@ (Random Variable)",
  "csv": "[任意] CSVデータ読込 (CSV Data Set Config)",
  "excel_row": "[任意] Excelデータ行取得 (JSR223 PreProcessor)",
  "access_token": "[任意] アクセストークン読込 (JSR223 PreProcessor)",
  "common_pre": "[任意] 共通前処理 (JSR223 PreProcessor)",
  "common_post": "[任意] 共通後処理 (JSR223 PostProcessor)",
  "throughput": "[任意] スループット制御 (Constant Throughput Timer)",
  "api_headers": "[任意] @@ID@@ ヘッダ (HTTP Header Manager)",
  "api_timer": "[任意] @@ID@@ 待機 @@MS@@ms (Constant Timer)",
  "api_pre": "[任意] @@ID@@ 前処理 (JSR223 PreProcessor)",
  "api_regex": "[任意] @@ID@@ 正規表現抽出 @@VAR@@ (Regular Expression Extractor)",
  "api_json": "[任意] @@ID@@ JSON抽出 @@VAR@@ (JSON Extractor)",
  "api_post": "[任意] @@ID@@ 後処理 (JSR223 PostProcessor)",
  "api_status": "[任意] @@ID@@ ステータス検証 @@CODES@@ (Response Assertion)",
  "api_text": "[任意] @@ID@@ 応答本文検証 (Response Assertion)",
  "api_duration": "[任意] @@ID@@ 応答時間検証 @@MS@@ms (Duration Assertion)",
  "debug": "[任意] 変数確認 (Debug Sampler)",
  "teardown_tg": "[任意] 事後処理 (tearDown Thread Group)",
  "teardown_end": "[任意] テスト終了処理 (JSR223 Sampler)",
  "view_tree": "[任意] 結果ツリー (View Results Tree)",
  "summary": "[任意] 集計レポート (Summary Report)",
  "aggregate": "[任意] 統計レポート (Aggregate Report)",
  "jtl": "[任意] JTL出力 全件 (Simple Data Writer)",
  "error_xml": "[任意] エラー詳細出力 XML (Simple Data Writer)",
  "result_log": "[任意] 結果ログ出力 (JSR223 Listener)"
 },
 "comments": {
  "test_plan": "jmeter_api_tool v1.0.0 で生成（@@DATE@@ / 入力: @@INPUT@@）\n名前が [必須] の要素は削除しないでください。[任意] の要素は不要であれば削除できます（各要素のコメント参照）。\n実行時の変更例: -JTHREADS=10 -JLOOPS=100 -JHOST=example.com",
  "udv": "[必須] 接続先と共通変数。値は ${__P(名前,既定値)} 形式のため、実行時に -J名前=値 で変更できます。\n各 API（HTTP リクエスト）が ${PROTOCOL} ${HOST} ${PORT} を参照しているため削除しないでください。",
  "udv_protocol": "プロトコル（http / https）。-J@@VAR@@=値 で変更",
  "udv_host": "ホスト名。-J@@VAR@@=値 で変更",
  "udv_port": "ポート番号（空欄 = http は 80 / https は 443）。-J@@VAR@@=値 で変更",
  "udv_user": "共通変数。-J@@VAR@@=値 で変更",
  "http_defaults": "[任意] 全 HTTP リクエストの初期値（文字コード・タイムアウト）。\nタイムアウトは -JCONNECT_TIMEOUT / -JRESPONSE_TIMEOUT（ミリ秒）で変更できます。削除するとタイムアウトは無制限になります。",
  "common_headers": "[任意] 全 API に付けるヘッダ。API 固有のヘッダは各 API の配下にあります（同じ名前の場合は API 側が優先）。",
  "cookie": "[任意] クッキーの自動管理（Set-Cookie を保存し、以降のリクエストで送信）。反復ごとにクリア = @@CLEAR@@。\nクッキーを使わない API だけの場合は削除できます。",
  "setup_tg": "[任意] テスト開始前に 1 回だけ実行されます。配下の処理が不要であれば、このスレッドグループごと削除できます。",
  "setup_start": "[任意] 実行条件（パラメータ欄）を jmeter.log に出力し、結果ログ用の集計値をリセットします。",
  "excel_loader": "[任意] Excel（.xlsx）を読み込みます（Apache POI 不要）。1 行目 = 変数名、2 行目以降 = データ。\nパラメータ: ファイルパス|シート名（-JEXCEL_FILE / -JEXCEL_SHEET で変更可。相対パスは JMX のフォルダ基準）。\n削除する場合は、メイン処理の「Excelデータ行取得」も削除してください。",
  "main_tg": "[必須] API を実行するスレッドグループ。\nスレッド数 = -JTHREADS / ランプアップ秒 = -JRAMP_UP / ループ回数 = -JLOOPS（-1 = 無限）@@DURATION@@",
  "main_tg_duration": " / 実行時間（秒）= -JDURATION",
  "counter": "[任意] 反復（ループ）ごとに連番を ${@@NAME@@} に設定します（@@SCOPE@@）。",
  "counter_scope_global": "全スレッドで一意",
  "counter_scope_thread": "スレッドごとに採番",
  "random": "[任意] 反復ごとに @@MIN@@ ～ @@MAX@@ の乱数を ${@@NAME@@} に設定します。",
  "csv": "[任意] CSV ファイルを 1 反復につき 1 行読み込み、列を変数に設定します（変数: @@VARS@@）。\nファイルは -JCSV_FILE で変更できます（相対パスは JMX のフォルダ基準）。ファイルが無いとスレッドが停止するため、使わない場合は削除してください。",
  "excel_row": "[任意] 「Excelデータ読込」で読み込んだデータから、反復ごとに次の 1 行を取り出して変数に設定します。",
  "access_token": "[任意] 全 API の呼び出し直前に、指定したフォルダ配下のファイルからアクセストークンを読み込み ${ACCESS_TOKEN} に設定します（ファイルが更新されると読み直します）。共通ヘッダの Authorization がこの変数を参照します。\nパラメータ: フォルダ|ファイル名（-JACCESS_TOKEN_DIR / -JACCESS_TOKEN_FILE で変更可。相対パスは JMX のフォルダ基準）。\nファイル名が *.token のように * を含む場合は、一致するファイルのうち更新日時が最も新しいファイルを使います。フォルダは ; で区切って複数書けます（存在する最初のフォルダを使用。例: C:\\tokens;/opt/tokens）。",
  "common_pre": "[任意] 全 API の呼び出し直前に毎回実行（Groovy）。${REQUEST_ID} ${TIMESTAMP} ${EPOCH_MS} を設定します。",
  "common_post": "[任意] 全 API の呼び出し直後に毎回実行（Groovy）。${LAST_STATUS} ${LAST_ELAPSED_MS} を設定します。",
  "throughput": "[任意] スループット制御。メイン処理全体で 1 分あたり -JTHROUGHPUT 件（既定 @@VALUE@@）になるよう待機を調整します。\n無効の場合は、使うときに右クリック → Enable で有効にしてください。",
  "sampler": "@@ID@@: @@METHOD@@ @@URL@@",
  "api_headers": "[任意] この API だけに付けるヘッダ。",
  "api_timer": "[任意] この API の呼び出し前に @@MS@@ ミリ秒待機します。",
  "api_pre": "[任意] この API の呼び出し直前に実行（Groovy）。",
  "api_post": "[任意] この API の呼び出し直後に実行（Groovy）。正規表現・JSON 抽出の後に実行されます。",
  "api_regex": "[任意] 応答本文から正規表現で抽出し ${@@VAR@@} に設定します（見つからない場合は NOT_FOUND）。",
  "api_json": "[任意] 応答 JSON から @@PATH@@ を抽出し ${@@VAR@@} に設定します（見つからない場合は NOT_FOUND）。",
  "api_status": "[任意] HTTP ステータスが @@CODES@@ であることを検証します。",
  "api_status_ignore": "（4xx / 5xx を期待するため「Ignore Status」を有効にしています）",
  "api_text": "[任意] 応答本文に指定した文字列がすべて含まれることを検証します。",
  "api_duration": "[任意] 応答時間が @@MS@@ ミリ秒以内であることを検証します。",
  "debug": "[任意] デバッグ用。全変数を結果ツリーに表示します。無効の場合は、使うときに右クリック → Enable で有効にしてください。",
  "teardown_tg": "[任意] 全スレッドの終了後に 1 回だけ実行されます。",
  "teardown_end": "[任意] 合計・成功・失敗件数を jmeter.log と結果ログに出力し、結果ログを閉じます。",
  "view_tree": "[任意] GUI でリクエスト / レスポンスを確認するためのリスナー。CLI 実行時は不要です（削除・無効化可）。",
  "summary": "[任意] GUI で集計結果を確認するためのリスナー。",
  "aggregate": "[任意] GUI で統計（パーセンタイル）を確認するためのリスナー。",
  "jtl": "[任意] 全リクエストの結果を JTL（CSV）に出力します。出力先: -JRESULT_DIR（既定 @@DIR@@、~/ は JMX のフォルダ）。\nrun コマンドや -l 指定時は同じ内容が result.jtl にも出力されるため、大量実行時は削除・無効化してかまいません。",
  "error_xml": "[任意] 失敗したリクエストだけを、リクエスト・レスポンス全文付きで XML に出力します（結果ツリーで読み込み可能）。",
  "result_log": "[任意] 1 リクエスト 1 行の結果ログ（api_result.log）を出力します。パラメータ: 出力フォルダ|本文出力(none/error/all)|最大文字数。\n大量実行時はファイル書き込みが負荷になるため、削除・無効化を推奨します。"
 },
 "scripts": {
  "excel_loader": "// ============================================================\n// [任意] Excelデータ読込（setUp Thread Group / JSR223 Sampler）\n//  テスト開始前に1回だけ実行され、Excel(.xlsx)の内容をメモリに読み込みます。\n//  ・Apache POI などの追加ライブラリは不要です（xlsx を ZIP + XML として直接解析）。\n//  ・1行目 = 列名（そのまま変数名になります）、2行目以降 = データ。空行は無視します。\n//  ・読み込んだデータは props 'EXCEL_DATA' に保存され、メイン処理の\n//    「Excelデータ行取得」が反復ごとに1行ずつ変数に設定します。\n//  ・パラメータ欄: ファイルパス|シート名（シート名が空の場合は先頭シート）\n//    相対パスは JMX ファイルのフォルダが基準です。-JEXCEL_FILE / -JEXCEL_SHEET で変更できます。\n//  ・数値・日付はセルに保存されている値（表示書式の適用前）になります。\n//    表示どおりの値を使いたい列は、Excel 側でセルの書式を「文字列」にしてください。\n// ============================================================\nimport java.util.concurrent.atomic.AtomicLong\nimport java.util.zip.ZipFile\nimport javax.xml.parsers.DocumentBuilderFactory\nimport org.apache.jmeter.services.FileServer\n\nSampleResult.setIgnore()   // この処理はレポート・JTL に含めない\n\nString[] params = (Parameters ?: '').split('\\\\|', -1)\nString path = params.length > 0 ? params[0].trim() : ''\nString sheetName = params.length > 1 ? params[1].trim() : ''\nprops.remove('EXCEL_DATA')\nprops.put('EXCEL_ROW_COUNTER', new AtomicLong(0))\nif (path.isEmpty()) {\n    log.warn('[Excel] ファイルパスが指定されていません')\n    return\n}\nFile file = new File(path)\nif (!file.isAbsolute()) {\n    file = new File(FileServer.getFileServer().getBaseDir(), path)\n}\nif (!file.isFile()) {\n    log.warn('[Excel] ファイルが見つかりません: ' + file.getAbsolutePath())\n    return\n}\n\ndef factory = DocumentBuilderFactory.newInstance()\nfactory.setNamespaceAware(true)\nfactory.setFeature('http://apache.org/xml/features/disallow-doctype-decl', true)\ndef builder = factory.newDocumentBuilder()\n\n// 子要素をローカル名で取得する\ndef children = { node, String name ->\n    def list = []\n    if (node == null) {\n        return list\n    }\n    def nodes = node.getChildNodes()\n    for (int i = 0; i < nodes.getLength(); i++) {\n        def n = nodes.item(i)\n        if (n.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE && n.getLocalName() == name) {\n            list.add(n)\n        }\n    }\n    return list\n}\n// <t> のテキストを連結する（ふりがな rPh は除外）\ndef textOf\ntextOf = { node ->\n    StringBuilder sb = new StringBuilder()\n    def nodes = node.getChildNodes()\n    for (int i = 0; i < nodes.getLength(); i++) {\n        def n = nodes.item(i)\n        if (n.getNodeType() != org.w3c.dom.Node.ELEMENT_NODE || n.getLocalName() == 'rPh') {\n            continue\n        }\n        sb.append(n.getLocalName() == 't' ? n.getTextContent() : textOf(n))\n    }\n    return sb.toString()\n}\n// セル参照 \"AB12\" → 列番号（0 始まり）\ndef columnIndex = { String ref ->\n    int n = 0\n    for (char ch : ref.toCharArray()) {\n        if (!Character.isLetter(ch)) {\n            break\n        }\n        n = n * 26 + (Character.toUpperCase(ch) - ('A' as char) + 1)\n    }\n    return n - 1\n}\n\ndef zip = new ZipFile(file)\ntry {\n    def parse = { String entryName ->\n        def entry = zip.getEntry(entryName)\n        if (entry == null) {\n            return null\n        }\n        def input = zip.getInputStream(entry)\n        try {\n            return builder.parse(input).getDocumentElement()\n        } finally {\n            input.close()\n        }\n    }\n    // 共有文字列\n    List<String> shared = []\n    children(parse('xl/sharedStrings.xml'), 'si').each { shared.add(textOf(it)) }\n    // シートの特定（workbook.xml → workbook.xml.rels）\n    Map<String, String> rels = [:]\n    children(parse('xl/_rels/workbook.xml.rels'), 'Relationship').each {\n        rels.put(it.getAttribute('Id'), it.getAttribute('Target'))\n    }\n    String target = null\n    String foundName = null\n    def sheetsNode = children(parse('xl/workbook.xml'), 'sheets')\n    for (def sheet : children(sheetsNode ? sheetsNode[0] : null, 'sheet')) {\n        if (sheetName.isEmpty() || sheet.getAttribute('name') == sheetName) {\n            def attrs = sheet.getAttributes()\n            for (int i = 0; i < attrs.getLength(); i++) {\n                def a = attrs.item(i)\n                if (a.getLocalName() == 'id' && a.getNamespaceURI() != null) {\n                    target = rels.get(a.getNodeValue())\n                }\n            }\n            foundName = sheet.getAttribute('name')\n            break\n        }\n    }\n    if (target == null) {\n        log.warn('[Excel] シートが見つかりません: ' + (sheetName ?: '(先頭シート)') + ' / ' + file.getName())\n        return\n    }\n    def sheetRoot = parse(target.startsWith('/') ? target.substring(1) : 'xl/' + target)\n    def sheetData = children(sheetRoot, 'sheetData')\n    // セルの読み取り\n    List<List<String>> table = []\n    for (def row : children(sheetData ? sheetData[0] : null, 'row')) {\n        Map<Integer, String> cells = [:]\n        int next = 0\n        for (def c : children(row, 'c')) {\n            String ref = c.getAttribute('r')\n            int col = ref ? columnIndex(ref) : next\n            next = col + 1\n            String type = c.getAttribute('t')\n            String value = ''\n            if (type == 'inlineStr') {\n                def inline = children(c, 'is')\n                value = inline ? textOf(inline[0]) : ''\n            } else {\n                def v = children(c, 'v')\n                String raw = v ? v[0].getTextContent() : ''\n                if (type == 's') {\n                    value = raw ? shared[Integer.parseInt(raw.trim())] : ''\n                } else if (type == 'b') {\n                    value = raw.trim() == '1' ? 'TRUE' : 'FALSE'\n                } else {\n                    value = raw\n                }\n            }\n            cells.put(col, value)\n        }\n        int size = cells.isEmpty() ? 0 : cells.keySet().max() + 1\n        List<String> values = []\n        for (int i = 0; i < size; i++) {\n            values.add(cells.containsKey(i) ? cells.get(i) : '')\n        }\n        if (values.any { it.trim() }) {\n            table.add(values)\n        }\n    }\n    if (table.isEmpty()) {\n        log.warn('[Excel] データがありません: ' + file.getName() + ' / ' + foundName)\n        return\n    }\n    List<String> headers = table[0].collect { it.trim() }\n    List<Map<String, String>> rows = []\n    for (int r = 1; r < table.size(); r++) {\n        Map<String, String> m = new LinkedHashMap<>()\n        for (int i = 0; i < headers.size(); i++) {\n            if (headers[i]) {\n                m.put(headers[i], i < table[r].size() ? table[r][i] : '')\n            }\n        }\n        rows.add(m)\n    }\n    props.put('EXCEL_DATA', Collections.unmodifiableList(rows))\n    log.info('[Excel] 読込完了: ' + file.getAbsolutePath() + ' / シート=' + foundName\n            + ' / 列=' + headers.findAll { it } + ' / データ=' + rows.size() + '行')\n} catch (Exception e) {\n    log.error('[Excel] 読込エラー: ' + file.getAbsolutePath(), e)\n} finally {\n    zip.close()\n}\n",
  "excel_row": "// ============================================================\n// [任意] Excelデータ行取得（JSR223 PreProcessor）\n//  「Excelデータ読込」（setUp）で読み込んだデータから、反復（ループ）ごとに\n//  次の1行を取り出し、1行目の列名を変数名として設定します。\n//    例: 列名 ITEM_CODE → 変数 ITEM_CODE、使用中の行番号 → 変数 EXCEL_ROW_NO\n//  ・全スレッドで共有し、上から順に払い出します。最終行の次は先頭に戻ります。\n//  ・同じ反復内の2つ目以降の API では同じ行の値を使います。\n// ============================================================\ndef rows = props.get('EXCEL_DATA')\nif (rows == null || rows.isEmpty()) {\n    return\n}\nString iteration = String.valueOf(vars.getIteration())\nif (iteration == vars.get('EXCEL_ROW_ITERATION')) {\n    return\n}\nvars.put('EXCEL_ROW_ITERATION', iteration)\ndef counter = props.get('EXCEL_ROW_COUNTER')\nlong n = counter != null ? counter.getAndIncrement() : 0L\nint index = (int) (n % rows.size())\nrows.get(index).each { k, v -> vars.put(k, v) }\nvars.put('EXCEL_ROW_NO', String.valueOf(index + 1))\n",
  "access_token": "// ============================================================\n// [任意] アクセストークン読込（JSR223 PreProcessor）\n//  メイン処理内のすべての API の呼び出し直前に毎回実行されます。\n//  指定したフォルダ配下のファイルからアクセストークンを読み込み、変数 ACCESS_TOKEN に設定します。\n//  共通ヘッダ（HTTP Header Manager）の Authorization（Bearer + 変数 ACCESS_TOKEN）で送信されます。\n//  ・パラメータ欄: フォルダ|ファイル名\n//    相対パスは JMX ファイルのフォルダ基準です。-JACCESS_TOKEN_DIR / -JACCESS_TOKEN_FILE で変更できます。\n//    区切りは \\ と / のどちらでもかまいません（Windows / Linux 共通）。フォルダを ; で区切って複数書くと、\n//    最初に存在するフォルダを使います（例: C:\\tokens;/opt/tokens）。\n//  ・ファイル名に * ? を含む場合（例: *.token）は、一致するファイルのうち更新日時が最も新しいファイルを使います\n//    （フォルダの確認は 1 秒に 1 回。新しいファイルを置くと自動で切り替わります）。\n//  ・ファイルの内容の前後の空白・改行（BOM）は除きます。\n//  ・ファイルの更新日時・サイズが変わったときだけ読み直します（実行中にトークンを差し替えできます）。\n// ============================================================\nimport java.nio.file.FileSystems\nimport org.apache.jmeter.services.FileServer\n\nString[] params = (Parameters ?: '').split('\\\\|', -1)\nString dir = params.length > 0 ? params[0].trim() : ''\nString name = params.length > 1 ? params[1].trim().replace('\\\\', '/') : ''\n// パスの解決: \\ は / として扱い、相対パス（~/ で始まるものを含む）は JMX のフォルダ基準\ndef resolve = { String path ->\n    String p = path.trim().replace('\\\\', '/')\n    File f = new File(p.startsWith('~/') ? p.substring(2) : p)\n    return f.isAbsolute() ? f : new File(FileServer.getFileServer().getBaseDir(), f.getPath())\n}\n// フォルダを ; で区切って複数書いた場合は、最初に存在するフォルダを使う\nList<File> dirs = dir.split(';').toList().findAll { it.trim() }.collect { resolve(it) }\nFile folder = dirs.find { it.isDirectory() } ?: (dirs ? dirs[0] : null)\nFile file = folder == null ? resolve(name) : new File(folder, name)\n// ファイル名に * ? がある場合: 一致するファイルのうち更新日時が最も新しいもの（同時刻は名前が後のもの）。\n// フォルダの確認は全スレッドで共有し、1 秒に 1 回だけ行う\nif (name.contains('*') || name.contains('?')) {\n    def latest = props.get('ACCESS_TOKEN_LATEST')\n    long now = System.currentTimeMillis()\n    if (latest == null || latest[0] != file.getPath() || now - latest[1] >= 1000) {\n        def matcher = FileSystems.getDefault().getPathMatcher('glob:' + file.getName())\n        File found = null\n        for (File f : (file.getParentFile()?.listFiles() ?: new File[0])) {\n            if (f.isFile() && matcher.matches(f.toPath().getFileName()) && (found == null || f.lastModified() > found.lastModified()\n                    || (f.lastModified() == found.lastModified() && f.getName() > found.getName()))) {\n                found = f\n            }\n        }\n        latest = [file.getPath(), now, found ?: file]\n        props.put('ACCESS_TOKEN_LATEST', latest)\n    }\n    file = latest[2]\n}\n// 読み込んだトークンは全スレッドで共有し、ファイルが変わったときだけ読み直す\nString stamp = file.getAbsolutePath() + '|' + file.lastModified() + '|' + file.length()\ndef cache = props.get('ACCESS_TOKEN_CACHE')\nif (cache == null || cache[0] != stamp) {\n    String token = ''\n    if (file.isFile()) {\n        token = file.getText('UTF-8').replace('﻿', '').trim()\n        if (token.isEmpty()) {\n            log.warn('[AccessToken] ファイルが空です: ' + file.getAbsolutePath())\n        } else {\n            log.info('[AccessToken] 読み込みました: ' + file.getAbsolutePath())\n        }\n    } else {\n        log.warn('[AccessToken] ファイルが見つかりません: ' + file.getAbsolutePath())\n    }\n    cache = [stamp, token]\n    props.put('ACCESS_TOKEN_CACHE', cache)\n}\nvars.put('ACCESS_TOKEN', cache[1])",
  "common_pre": "// ============================================================\n// [任意] 共通前処理（JSR223 PreProcessor）\n//  メイン処理内のすべての API の呼び出し直前に毎回実行されます。\n//  API 呼び出し前にパラメータを生成・加工する処理をここに書きます。\n//  ※ JSR223 内では JMeter の変数参照（ドル記号＋波括弧）は使わず、\n//     vars.get('変数名') / vars.put('変数名', 値) を使います。\n// ============================================================\nimport java.time.ZonedDateTime\nimport java.time.format.DateTimeFormatter\n\n// リクエストごとの一意な ID とタイムスタンプ → 変数 REQUEST_ID / TIMESTAMP / EPOCH_MS\nvars.put('REQUEST_ID', UUID.randomUUID().toString())\nvars.put('TIMESTAMP', ZonedDateTime.now().format(DateTimeFormatter.ISO_OFFSET_DATE_TIME))\nvars.put('EPOCH_MS', String.valueOf(System.currentTimeMillis()))\n\n// 加工の例（必要に応じてコメントを外して使用）\n// String userId = vars.get('USER_ID') ?: ''\n// vars.put('USER_ID_UPPER', userId.toUpperCase())\n// vars.put('TODAY', new Date().format('yyyyMMdd'))\n",
  "common_post": "// ============================================================\n// [任意] 共通後処理（JSR223 PostProcessor）\n//  メイン処理内のすべての API の呼び出し直後に毎回実行されます。\n//  応答からパラメータを取得し、次の API で使う変数に設定する処理をここに書きます。\n//  prev = 直前の API の実行結果（SampleResult）\n//  ※ 検証（Assertion）より前に実行されるため、prev.isSuccessful() は HTTP レベルの成否です。\n// ============================================================\nvars.put('LAST_STATUS', prev.getResponseCode())\nvars.put('LAST_ELAPSED_MS', String.valueOf(prev.getTime()))\n\n// 応答 JSON から値を取り出す例（必要に応じてコメントを外して使用）\n// if ((prev.getContentType() ?: '').contains('json')) {\n//     def json = new groovy.json.JsonSlurper().parseText(prev.getResponseDataAsString())\n//     vars.put('RESULT_CODE', String.valueOf(json?.resultCode))\n// }\n// 応答ヘッダから値を取り出す例\n// def m = prev.getResponseHeaders() =~ /(?im)^X-Request-Id:\\s*(.+)$/\n// if (m.find()) {\n//     vars.put('RESPONSE_REQUEST_ID', m.group(1).trim())\n// }\n",
  "setup_start": "// ============================================================\n// [任意] テスト開始処理（setUp Thread Group / JSR223 Sampler）\n//  テスト開始時に1回だけ実行されます。\n//  ・実行条件（パラメータ欄）を jmeter.log に出力します。\n//  ・前回の実行（GUI で連続実行した場合）の集計値・結果ログをリセットします。\n//  削除しても他の要素の動作には影響しません。\n// ============================================================\nSampleResult.setIgnore()   // この処理はレポート・JTL に含めない\n\ndef oldWriter = props.remove('API_RESULT_LOG_WRITER')\nif (oldWriter != null) {\n    try {\n        oldWriter.close()\n    } catch (Exception ignored) {\n    }\n}\n['API_RESULT_TOTAL', 'API_RESULT_OK', 'API_RESULT_NG'].each { props.remove(it) }\nprops.put('API_TEST_START_MS', String.valueOf(System.currentTimeMillis()))\nlog.info('==== テスト開始: ' + (Parameters ?: ''))\n",
  "teardown_end": "// ============================================================\n// [任意] テスト終了処理（tearDown Thread Group / JSR223 Sampler）\n//  すべてのスレッドの終了後に1回だけ実行されます。\n//  ・合計 / 成功 / 失敗件数と経過時間を jmeter.log と結果ログに出力します。\n//    （件数は「結果ログ出力」が有効な場合のみ集計されます）\n//  ・結果ログファイルを閉じ、読み込んだ Excel データを解放します。\n// ============================================================\nSampleResult.setIgnore()   // この処理はレポート・JTL に含めない\n\ndef total = props.get('API_RESULT_TOTAL')\nString counts = total == null ? '件数=集計なし（結果ログ出力が無効）'\n        : '合計=' + total.get() + ' 成功=' + (props.get('API_RESULT_OK')?.get() ?: 0) + ' 失敗=' + (props.get('API_RESULT_NG')?.get() ?: 0)\nString start = props.get('API_TEST_START_MS')\nString elapsed = start ? ' 経過秒=' + (System.currentTimeMillis() - Long.parseLong(start)).intdiv(1000) : ''\nString summary = '==== テスト終了: ' + counts + elapsed\nlog.info(summary)\n\ndef writer = props.remove('API_RESULT_LOG_WRITER')\nif (writer != null) {\n    synchronized (writer) {\n        writer.println('# ' + summary)\n        writer.close()\n    }\n}\nprops.remove('EXCEL_DATA')\nprops.remove('EXCEL_ROW_COUNTER')\n",
  "result_listener": "// ============================================================\n// [任意] 結果ログ出力（JSR223 Listener）\n//  すべての API の実行結果を 1リクエスト1行（タブ区切り）で api_result.log に出力します。\n//  パラメータ欄: 出力フォルダ|本文出力|最大文字数\n//   ・出力フォルダ: -JRESULT_DIR で変更（~/ は JMX のフォルダ。run コマンドは自動設定）\n//   ・本文出力  : none = 出力しない / error = 失敗時のみ / all = 全件（-JLOG_RESPONSE）\n//   ・最大文字数: リクエスト・レスポンス本文の最大文字数（-JLOG_MAX_CHARS）\n//  ※ 大量実行時はファイル書き込みが負荷になるため、無効化または none を推奨します。\n// ============================================================\nimport java.text.SimpleDateFormat\nimport java.util.concurrent.atomic.AtomicLong\nimport java.util.function.Function\nimport org.apache.jmeter.services.FileServer\n\ndef res = sampleResult\nif (res == null) {\n    return\n}\nString[] p = (Parameters ?: '').split('\\\\|', -1)\nString dirText = (p.length > 0 && p[0].trim()) ? p[0].trim() : '~/result'\nString mode = (p.length > 1 && p[1].trim()) ? p[1].trim().toLowerCase() : 'error'\nint maxChars = (p.length > 2 && p[2].trim().isInteger()) ? p[2].trim().toInteger() : 1000\n\n// 件数の集計（テスト終了処理で出力）\ndef count = { String key -> props.computeIfAbsent(key, { k -> new AtomicLong(0) } as Function).incrementAndGet() }\ncount('API_RESULT_TOTAL')\ncount(res.isSuccessful() ? 'API_RESULT_OK' : 'API_RESULT_NG')\n\n// 出力ファイル（最初の1件で作成し、全スレッドで共有）\ndef writer = props.computeIfAbsent('API_RESULT_LOG_WRITER', { k ->\n    File dir = dirText.startsWith('~/') ? new File(FileServer.getFileServer().getBaseDir(), dirText.substring(2)) : new File(dirText)\n    dir.mkdirs()\n    File file = new File(dir, 'api_result.log')\n    boolean isNew = !file.exists() || file.length() == 0\n    def w = new PrintWriter(new OutputStreamWriter(new FileOutputStream(file, true), 'UTF-8'), true)\n    if (isNew) {\n        w.println(['日時', '結果', 'ステータス', '応答時間ms', '受信バイト', 'ラベル', 'リクエスト', 'スレッド', 'メッセージ'].join('\\t'))\n    }\n    log.info('[結果ログ] 出力先: ' + file.getAbsolutePath())\n    return w\n} as Function)\n\n// 改行・タブを除去し、最大文字数で切り詰める\ndef oneLine = { String s, int limit ->\n    if (s == null) {\n        return ''\n    }\n    String t = s.replace('\\r', '').replace('\\n', '\\\\n').replace('\\t', ' ')\n    return (limit > 0 && t.length() > limit) ? t.substring(0, limit) + '...(全' + t.length() + '文字)' : t\n}\nString request = ''\nif (res instanceof org.apache.jmeter.protocol.http.sampler.HTTPSampleResult) {\n    request = (res.getHTTPMethod() ?: '') + ' ' + (res.getUrlAsString() ?: '')\n}\nString message = res.getResponseMessage() ?: ''\ndef failures = res.getAssertionResults().findAll { it.isFailure() || it.isError() }\nif (failures) {\n    message += ' / ' + failures.collect { it.getName() + ': ' + it.getFailureMessage() }.join(' / ')\n}\nString time = new SimpleDateFormat('yyyy-MM-dd HH:mm:ss.SSS').format(new Date(res.getStartTime()))\nString line = [time, res.isSuccessful() ? 'OK' : 'NG', res.getResponseCode(), res.getTime(), res.getBytesAsLong(),\n               res.getSampleLabel(), request.trim(), res.getThreadName(), oneLine(message, 1000)].join('\\t')\nsynchronized (writer) {\n    writer.println(line)\n    if (mode == 'all' || (mode == 'error' && !res.isSuccessful())) {\n        writer.println('\\t[REQUEST] ' + oneLine(res.getSamplerData(), maxChars))\n        writer.println('\\t[RESPONSE] ' + oneLine(res.getResponseDataAsString(), maxChars))\n    }\n}\nif (!res.isSuccessful()) {\n    log.warn('[API NG] ' + res.getSampleLabel() + ' status=' + res.getResponseCode() + ' ' + oneLine(message, 500))\n}\n"
 },
 "template": {
  "api_sheet": "API一覧",
  "settings_sheet": "設定",
  "guide_sheet": "記入方法",
  "api_title": "API定義一覧",
  "api_note": "1行 = 1API（上から順に実行）。複数の値はセル内改行（Alt + Enter）で区切ります。4行目（No が # で始まる行）は記入説明のため読み込まれません。詳しくは「記入方法」シートを参照してください。",
  "settings_title": "全体設定",
  "settings_note": "「設定値」列を変更してください（空欄は既定値）。ON = 生成して有効 / OFF = 生成しない / DISABLE = 生成するが無効化。説明に -J名前 とある項目は、実行時にも -J名前=値 で変更できます。",
  "settings_headers": [
   "設定キー",
   "設定項目",
   "設定値",
   "既定値",
   "説明"
  ],
  "settings_widths": [
   28,
   24,
   40,
   16,
   74
  ],
  "empty_default": "（空欄）",
  "blank_rows": 20,
  "validation_rows": 300
 },
 "styles": {
  "fonts": [
   {
    "sz": 10,
    "color": "FF000000"
   },
   {
    "sz": 10,
    "b": true,
    "color": "FF000000"
   },
   {
    "sz": 16,
    "b": true,
    "color": "FF1A1A1A"
   },
   {
    "sz": 11,
    "b": true,
    "color": "FFFFFFFF"
   },
   {
    "sz": 10,
    "b": true,
    "color": "FFFFFFFF"
   },
   {
    "sz": 9,
    "color": "FF595959"
   },
   {
    "sz": 10,
    "color": "FF404040"
   },
   {
    "sz": 20,
    "b": true,
    "color": "FF1A1A1A"
   },
   {
    "sz": 10,
    "u": true,
    "color": "FF262626"
   },
   {
    "sz": 11,
    "b": true,
    "color": "FF262626"
   }
  ],
  "fills": [
   "",
   "",
   "FF262626",
   "FF404040",
   "FF595959",
   "FFD9D9D9",
   "FFF2F2F2",
   "FFF8F8F8",
   "FFEDEDED",
   "FFFFFFFF"
  ],
  "borders": [
   {},
   {
    "left": [
     "thin",
     "FFA6A6A6"
    ],
    "right": [
     "thin",
     "FFA6A6A6"
    ],
    "top": [
     "thin",
     "FFA6A6A6"
    ],
    "bottom": [
     "thin",
     "FFA6A6A6"
    ]
   },
   {
    "left": [
     "thin",
     "FF7F7F7F"
    ],
    "right": [
     "thin",
     "FF7F7F7F"
    ],
    "top": [
     "thin",
     "FF7F7F7F"
    ],
    "bottom": [
     "thin",
     "FF7F7F7F"
    ]
   },
   {
    "bottom": [
     "medium",
     "FF1A1A1A"
    ]
   },
   {
    "bottom": [
     "thin",
     "FF7F7F7F"
    ]
   },
   {
    "left": [
     "thin",
     "FFBFBFBF"
    ],
    "right": [
     "thin",
     "FFBFBFBF"
    ],
    "top": [
     "thin",
     "FFBFBFBF"
    ]
   },
   {
    "left": [
     "thin",
     "FFBFBFBF"
    ],
    "right": [
     "thin",
     "FFBFBFBF"
    ]
   },
   {
    "left": [
     "thin",
     "FFBFBFBF"
    ],
    "right": [
     "thin",
     "FFBFBFBF"
    ],
    "bottom": [
     "thin",
     "FFBFBFBF"
    ]
   },
   {
    "left": [
     "thin",
     "FFBFBFBF"
    ],
    "right": [
     "thin",
     "FFBFBFBF"
    ],
    "top": [
     "thin",
     "FFBFBFBF"
    ],
    "bottom": [
     "thin",
     "FFBFBFBF"
    ]
   },
   {
    "left": [
     "medium",
     "FF595959"
    ]
   },
   {
    "left": [
     "thin",
     "FF595959"
    ],
    "right": [
     "thin",
     "FF595959"
    ],
    "top": [
     "thin",
     "FF595959"
    ],
    "bottom": [
     "thin",
     "FF595959"
    ]
   }
  ],
  "xfs": [
   [
    "default",
    0,
    0,
    0,
    "",
    "",
    false,
    0,
    0
   ],
   [
    "title",
    2,
    0,
    3,
    "",
    "center",
    false,
    0,
    0
   ],
   [
    "cover_title",
    7,
    0,
    0,
    "",
    "center",
    false,
    0,
    0
   ],
   [
    "cover_sub",
    6,
    0,
    0,
    "",
    "center",
    false,
    0,
    0
   ],
   [
    "h3",
    3,
    3,
    0,
    "",
    "center",
    false,
    1,
    0
   ],
   [
    "h4",
    9,
    0,
    4,
    "",
    "center",
    false,
    0,
    0
   ],
   [
    "p",
    0,
    0,
    0,
    "",
    "top",
    true,
    0,
    0
   ],
   [
    "li",
    0,
    0,
    0,
    "",
    "top",
    true,
    1,
    0
   ],
   [
    "quote",
    6,
    6,
    9,
    "",
    "top",
    true,
    1,
    0
   ],
   [
    "code_first",
    0,
    6,
    5,
    "",
    "top",
    true,
    1,
    0
   ],
   [
    "code_mid",
    0,
    6,
    6,
    "",
    "top",
    true,
    1,
    0
   ],
   [
    "code_last",
    0,
    6,
    7,
    "",
    "top",
    true,
    1,
    0
   ],
   [
    "code_single",
    0,
    6,
    8,
    "",
    "top",
    true,
    1,
    0
   ],
   [
    "th",
    4,
    4,
    2,
    "",
    "center",
    true,
    0,
    0
   ],
   [
    "td",
    0,
    0,
    1,
    "",
    "top",
    true,
    0,
    0
   ],
   [
    "td_alt",
    0,
    7,
    1,
    "",
    "top",
    true,
    0,
    0
   ],
   [
    "link",
    8,
    0,
    0,
    "",
    "center",
    false,
    0,
    0
   ],
   [
    "api_hdr",
    4,
    3,
    2,
    "center",
    "center",
    true,
    0,
    0
   ],
   [
    "api_hdr_req",
    4,
    2,
    2,
    "center",
    "center",
    true,
    0,
    0
   ],
   [
    "hint",
    5,
    8,
    1,
    "",
    "top",
    true,
    0,
    49
   ],
   [
    "cell",
    0,
    0,
    1,
    "",
    "top",
    true,
    0,
    49
   ],
   [
    "cat",
    1,
    5,
    1,
    "",
    "center",
    false,
    0,
    0
   ],
   [
    "key",
    0,
    6,
    1,
    "",
    "top",
    true,
    0,
    0
   ],
   [
    "input",
    0,
    9,
    10,
    "",
    "top",
    true,
    0,
    49
   ],
   [
    "note",
    5,
    0,
    0,
    "",
    "center",
    true,
    0,
    0
   ],
   [
    "data_hdr",
    4,
    4,
    2,
    "center",
    "center",
    false,
    0,
    0
   ],
   [
    "data_cell",
    0,
    0,
    1,
    "",
    "center",
    false,
    0,
    49
   ],
   [
    "td_center",
    0,
    0,
    1,
    "center",
    "top",
    true,
    0,
    0
   ],
   [
    "toc_no",
    0,
    0,
    1,
    "center",
    "center",
    false,
    0,
    0
   ],
   [
    "toc_link",
    8,
    0,
    1,
    "",
    "center",
    false,
    0,
    0
   ],
   [
    "toc_td",
    0,
    0,
    1,
    "",
    "center",
    true,
    0,
    0
   ]
  ]
 },
 "xlsx": {
  "font_name": "Meiryo UI",
  "line_h10": 144,
  "pad_h10": 36,
  "max_h10": 4090,
  "units": {
   "upper": 22,
   "digit": 20,
   "lower": 17,
   "punct": 15,
   "latin": 20,
   "hankaku": 17,
   "kana": 27,
   "wide": 33
  },
  "min_total_width": 110,
  "manual_file": "JMeter_Usage_Guide",
  "cover_sheet": "表紙",
  "toc_title": "目次",
  "toc_headers": [
   "No",
   "シート",
   "内容"
  ],
  "cover_widths": [
   6,
   26,
   82
  ]
 },
 "manual_md": "# JMeter APIテスト 利用方法\n\n> 対象: Apache JMeter 5.6.3 ／ RHEL 9.8 ／ Windows 11 ／ ツール: jmeter_api_tool v1.0.0（@@DATE@@ 出力）\n\n本書は、REST API の定義（URL・HTTPメソッド・クエリパラメータ・ヘッダ・リクエストボディ）を Excel または CSV に記入し、JMeter のテスト計画ファイル（JMX）を自動生成して実行し、結果を確認するまでの手順をまとめたものです。\n\n## 1.概要\n<!-- summary: ツールでできること・ファイル構成・作業の流れ -->\n\n### このツールでできること\n- Excel / CSV に記入した API 定義から、JMeter 5.6.3 用の JMX ファイルを自動生成します。\n- 生成される JMX には、API テストでよく使う要素（待機タイマ、連番・乱数、CSV / Excel データ読込、クッキー、前処理・後処理、正規表現 / JSON 抽出、応答の検証、結果出力）があらかじめ組み込まれています。\n- 既定は「1スレッド × 1回」の実行です。実行時のプロパティ指定（例: -JTHREADS=10 -JLOOPS=100）だけで、大量件数の負荷試験に切り替えられます。\n- 実行結果は JTL（CSV）、HTML レポート、jmeter.log、結果ログ（api_result.log）、エラー詳細（error_detail.xml）に出力されます。\n- 生成される要素の名前には [必須] / [任意] を付けています。[任意] の要素は JMeter GUI で右クリック →「削除」するだけで、他の要素に影響を与えずに取り除けます。\n\n### ファイル構成\n<!-- cols: 36,74 -->\n| パス | 内容 |\n|---|---|\n| jmeter_api_tool.sh | RHEL 9.8 用スクリプト（bash + python3 標準ライブラリのみ使用） |\n| jmeter_api_tool.ps1 | Windows 11 用スクリプト（Windows PowerShell 5.1 / PowerShell 7 で動作） |\n| jmeter_api_tool.bat | Windows 用の起動ファイル（実行ポリシーを一時的に回避して .ps1 を起動） |\n| input/api_definition.xlsx | 入力テンプレート（Excel）：「API一覧」「設定」「記入方法」シート |\n| input/api_definition.csv | 入力テンプレート（CSV）：API 一覧 |\n| input/api_definition_settings.csv | 入力テンプレート（CSV）：設定 |\n| input/openapi_config.csv | OpenAPI 定義（.json）から生成する場合の設定（接続先・HTTP ヘッダ・アクセストークン・定義の変数 ${public_base_url} / ${name_prefix} の値）。generate の -c / -OpenApiConfig で指定 |\n| output/api_test.jmx | 生成された JMX（サンプル） |\n| output/data/testdata.csv | CSV データ読込のサンプルデータ |\n| output/data/testdata.xlsx | Excel データ読込のサンプルデータ |\n| docs/JMeter_Usage_Guide.xlsx / .md | 本書（manual コマンドで出力） |\n| result/JMX名_日時/ | run コマンドの実行結果（JTL・ログ・HTML レポート） |\n\n### 作業の流れ\n1. テンプレートを作成する（init コマンド）\n2. API 定義と設定を記入する（Excel または CSV）\n3. JMX を生成する（generate コマンド）\n4. JMeter GUI で内容を確認し、不要な [任意] 要素を削除・調整する\n5. CLI モードで実行する（run コマンド）\n6. 結果（HTML レポート・JTL・ログ）を確認する\n\n> ファイル名・フォルダ名はすべて英数字にしています。Windows と Linux の間でファイルを転送したときの文字化けを防ぐためです。\n\n## 2.環境構築\n<!-- summary: Java・JMeter 5.6.3 のインストールと動作確認（RHEL 9.8 / Windows 11） -->\n\n### 必要なソフトウェア\n<!-- cols: 18,30,30,40 -->\n| ソフトウェア | RHEL 9.8 | Windows 11 | 備考 |\n|---|---|---|---|\n| Java | java-17-openjdk（dnf） | JDK 17（Eclipse Temurin 等） | JMeter 5.6.3 は Java 8 以上で動作。17 を推奨 |\n| Apache JMeter | 5.6.3（.tgz） | 5.6.3（.zip） | https://archive.apache.org/dist/jmeter/binaries/ |\n| スクリプト実行環境 | bash・python3（OS 標準） | Windows PowerShell 5.1（OS 標準） | 追加のモジュール・ライブラリは不要 |\n| Excel | 不要 | 不要（記入・閲覧には推奨） | xlsx の読み書きはスクリプト自身が行います |\n\n### RHEL 9.8 へのインストール\n```bash\n# Java 17（GUI を使わないサーバーなら java-17-openjdk-headless でも可）\nsudo dnf install -y java-17-openjdk\njava -version\n\n# JMeter 5.6.3（ダウンロード・チェックサム確認・展開）\ncd /opt\nsudo curl -LO https://archive.apache.org/dist/jmeter/binaries/apache-jmeter-5.6.3.tgz\nsudo curl -LO https://archive.apache.org/dist/jmeter/binaries/apache-jmeter-5.6.3.tgz.sha512\nsha512sum -c apache-jmeter-5.6.3.tgz.sha512\nsudo tar xzf apache-jmeter-5.6.3.tgz\n\n# 環境変数（~/.bashrc に追記）\necho 'export JMETER_HOME=/opt/apache-jmeter-5.6.3' >> ~/.bashrc\necho 'export PATH=$JMETER_HOME/bin:$PATH' >> ~/.bashrc\nsource ~/.bashrc\njmeter -v\n```\n- python3 は RHEL 9 のパッケージ管理（dnf）が使用しているため、通常は導入済みです（確認: `python3 --version`）。無い場合は `sudo dnf install -y python3` で導入します。\n- ツールのファイルを Windows から転送して改行コードが CRLF になった場合は、`sed -i 's/\\r$//' jmeter_api_tool.sh` で LF に戻してください。\n- GUI を使う場合に日本語が □ で表示されるときは、日本語フォントを導入します（例: `sudo dnf install -y google-noto-sans-cjk-ttc-fonts`）。\n\n### Windows 11 へのインストール\n1. JDK 17 をインストールします（例: Eclipse Temurin 17 の .msi。インストール時に「JAVA_HOME 変数の設定」と「PATH への追加」を有効にします）。\n2. https://archive.apache.org/dist/jmeter/binaries/apache-jmeter-5.6.3.zip をダウンロードし、ファイルのプロパティで「許可する（ブロック解除）」にチェックしてから、C:\\apache-jmeter-5.6.3 などスペースや日本語を含まないフォルダに展開します。\n3. 環境変数 JMETER_HOME に展開先のフォルダを設定します（下記のコマンド、または「システムの詳細設定」→「環境変数」）。\n4. PowerShell を開き直して動作を確認します。\n\n```powershell\n[Environment]::SetEnvironmentVariable('JMETER_HOME', 'C:\\apache-jmeter-5.6.3', 'User')\n# PowerShell を開き直してから確認\njava -version\n& \"$env:JMETER_HOME\\bin\\jmeter.bat\" -v\n```\n\n### ツールの配置と動作確認\nツール一式を任意のフォルダ（例: RHEL は ~/JMeter_API_Generate、Windows は C:\\work\\JMeter_API_Generate）に配置し、ヘルプが表示されることを確認します。\n\n```bash\n# RHEL 9.8\ncd ~/JMeter_API_Generate\nchmod +x jmeter_api_tool.sh\n./jmeter_api_tool.sh help\n```\n\n```powershell\n# Windows 11（PowerShell またはコマンドプロンプト）\ncd C:\\work\\JMeter_API_Generate\n.\\jmeter_api_tool.bat help\n```\n- jmeter_api_tool.bat は PowerShell の実行ポリシー（既定では .ps1 の実行が禁止）を、その実行に限って回避して起動します。PowerShell から直接実行する場合は `powershell -ExecutionPolicy Bypass -File .\\jmeter_api_tool.ps1 help` としてください。\n- インターネットから取得した ZIP を展開した場合は、事前に `Get-ChildItem -Recurse | Unblock-File` でブロックを解除しておくと確実です。\n\n## 3.JMX生成\n<!-- summary: テンプレート作成から JMX 生成・GUI での確認までの手順 -->\n\n### 手順1：テンプレートを作成する\n```bash\n./jmeter_api_tool.sh init          # RHEL 9.8\n```\n```powershell\n.\\jmeter_api_tool.bat init         # Windows 11\n```\ninput フォルダに次の 4 ファイルが作成されます（既存ファイルは上書きしません。上書きする場合は -f / -Force を付けます）。\n- api_definition.xlsx（Excel 形式のテンプレート）\n- api_definition.csv（CSV 形式：API 一覧）\n- api_definition_settings.csv（CSV 形式：設定）\n- openapi_config.csv（CSV 形式：OpenAPI 定義から生成する場合の接続先・HTTP ヘッダ・アクセストークン・定義の変数 ${public_base_url} / ${name_prefix} の値。generate の -c / -OpenApiConfig で指定）\n\n### 手順2：API 定義を記入する\n- Excel 形式：「API一覧」シートに 1行 = 1API で記入し、「設定」シートの「設定値」列で全体設定を変更します。\n- CSV 形式：api_definition.csv に API を、api_definition_settings.csv に設定を記入します。\n- テンプレートには記入例として、postman-echo.com（送信した内容をそのまま返す公開サービス）を呼び出す 4 件の API が入っています。自分の API に書き換えてください。\n- 入力項目の詳細は「4.入力項目_設定」「5.入力項目_API」シートを参照してください。\n\n> CSV を Excel で保存する場合は「CSV UTF-8（コンマ区切り）」を選んでください。Shift_JIS の「CSV（コンマ区切り）」で保存したファイルも自動判別して読み込めます。\n\n### 手順3：JMX を生成する\n```bash\n# RHEL 9.8\n./jmeter_api_tool.sh generate -i input/api_definition.xlsx\n./jmeter_api_tool.sh generate -i input/api_definition.csv -s input/api_definition_settings.csv\n```\n```powershell\n# Windows 11\n.\\jmeter_api_tool.bat generate -i input\\api_definition.xlsx\n.\\jmeter_api_tool.bat generate -i input\\api_definition.csv -s input\\api_definition_settings.csv\n```\n\n<!-- cols: 22,30,66 -->\n| sh のオプション | PowerShell のオプション | 説明 |\n|---|---|---|\n| -i ファイル | -InputFile（-i） | 入力ファイル（.xlsx / .csv）。必須 |\n| -s ファイル | -SettingsFile（-s） | 設定 CSV。省略時は「入力ファイル名_settings.csv」を探し、無ければ既定値を使用（Excel 入力時は「設定」シートを使用） |\n| -o フォルダ | -OutDir（-o） | 出力先フォルダ（既定: output） |\n| -n ファイル名 | -JmxName（-n） | 出力する JMX のファイル名（例: order_test.jmx）。設定「出力JMXファイル名」（output_file）より優先します |\n| --split | -Split | API ごとに別々の JMX を出力する |\n| -f | -Force（-f） | サンプルデータ（data/testdata.csv / .xlsx）が既にあっても上書きする |\n\n### 生成されるファイル\n<!-- cols: 34,84 -->\n| ファイル | 説明 |\n|---|---|\n| output/出力JMXファイル名 | 生成された JMX。同名のファイルがある場合は「.日時.bak」を付けて退避してから作成します |\n| output/data/testdata.csv | CSV データ読込用のサンプル（ファイルが無い場合のみ作成） |\n| output/data/testdata.xlsx | Excel データ読込用のサンプル（ファイルが無い場合のみ作成） |\n\n### 手順4：JMeter GUI で確認・調整する\n- RHEL（デスクトップ環境あり）：`jmeter -t output/api_test.jmx`\n- Windows：`%JMETER_HOME%\\bin\\jmeter.bat` をダブルクリックし、「File」→「Open」で JMX を選択\n- 不要な [任意] 要素は右クリック →「Remove」で削除、一時的に外す場合は右クリック →「Disable」で無効化できます。\n- GUI モードは動作確認・デバッグ用です。負荷試験は必ず CLI モード（run コマンド）で実行してください。\n\n## 4.入力項目_設定\n<!-- summary: 「設定」シート（設定 CSV）の全項目と既定値 -->\n「設定」シート（CSV の場合は api_definition_settings.csv）の項目です。設定値が空欄の項目は既定値が使われます。\n\n- ON / OFF / DISABLE の項目は、ON = 生成して有効、OFF = 生成しない、DISABLE = 生成するが無効化（GUI で右クリック →「Enable」ですぐ使える）を表します。\n- 説明欄の「-J名前」は、実行時に -J名前=値（run コマンドでは -J 名前=値 / -Prop 名前=値）で値を変更できることを表します。\n\n<!-- @SETTINGS_TABLE@ -->\n\n## 5.入力項目_API\n<!-- summary: 「API一覧」シート（API CSV）の列と記入ルール -->\n「API一覧」シート（CSV の場合は api_definition.csv）の列です。1行 = 1API で、上から順に実行されます。\n\n<!-- @API_COLUMNS_TABLE@ -->\n\n### 記入のルール\n- 複数の値を書く列（クエリパラメータ、ヘッダ、正規表現抽出など）は、セル内改行（Excel では Alt + Enter）で区切ります。クエリパラメータと form 形式のボディは & 区切りでも記入できます。\n- 先頭の列（No）が # で始まる行はコメントとして読み飛ばします（テンプレートの 4 行目の記入説明など）。\n- URL・パラメータ・ヘッダ・ボディには ${変数名} で JMeter の変数・関数を埋め込めます（例: `${USER_ID}`、`${SEQ}`、`${__UUID()}`）。\n- 「有効」を N にした API は、生成はされますが無効化された状態になります。\n- GET / DELETE のクエリパラメータは HTTP リクエストの「Parameters」タブに設定されます。POST / PUT / PATCH の場合は URL のパスに付与されます（値は UTF-8 で URL エンコード。${...} の部分はそのまま）。\n- 既に URL エンコード済みの値（%E3%83%86 など）は、二重にエンコードしないよう自動で判定します。\n- GET / DELETE ではリクエストボディは使用できません（指定した場合は警告を表示して無視します）。\n\n### ボディ形式（auto の場合の自動判定）\n<!-- cols: 26,16,76 -->\n| 本文の内容 | 判定 | 送信方法 |\n|---|---|---|\n| { または [ で始まる | json | 本文をそのまま送信。Content-Type が未指定なら application/json を付与 |\n| < で始まる | raw（XML） | 本文をそのまま送信。Content-Type が未指定なら application/xml を付与 |\n| 名前=値 の形式（改行 / & 区切り） | form | パラメータとして送信（application/x-www-form-urlencoded） |\n| 上記以外 | raw | 本文をそのまま送信（Content-Type はヘッダ列で指定） |\n\n## 6.JMX構成\n<!-- summary: 生成される JMX の要素一覧と、削除してよいかの判断基準 -->\n生成される JMX のツリー構成です（記入例を生成した場合）。名前が [必須] の要素は削除しないでください。[任意] の要素は、他の要素に影響を与えずに削除できます。\n\n```text\nテストプラン\n├─ [必須] 共通変数 (User Defined Variables)\n├─ [任意] HTTPリクエスト初期値 (HTTP Request Defaults)\n├─ [任意] 共通ヘッダ (HTTP Header Manager)\n├─ [任意] クッキー管理 (HTTP Cookie Manager)\n├─ [任意] 事前処理 (setUp Thread Group)\n│   ├─ [任意] テスト開始処理 (JSR223 Sampler)\n│   └─ [任意] Excelデータ読込 (JSR223 Sampler)\n├─ [必須] メイン処理 (Thread Group)\n│   ├─ [任意] 連番 SEQ (Counter)\n│   ├─ [任意] 乱数 RANDOM_NO (Random Variable)\n│   ├─ [任意] CSVデータ読込 (CSV Data Set Config)\n│   ├─ [任意] Excelデータ行取得 (JSR223 PreProcessor)\n│   ├─ [任意] 共通前処理 (JSR223 PreProcessor)\n│   ├─ [任意] 共通後処理 (JSR223 PostProcessor)\n│   ├─ [任意] スループット制御 (Constant Throughput Timer) ※既定は無効\n│   ├─ API01_ログイン(POST JSON)  ← HTTP リクエスト（1API = 1つ）\n│   │   ├─ [任意] API01 ヘッダ (HTTP Header Manager)\n│   │   ├─ [任意] API01 待機 500ms (Constant Timer)\n│   │   ├─ [任意] API01 前処理 (JSR223 PreProcessor)\n│   │   ├─ [任意] API01 正規表現抽出 ECHO_URL (Regular Expression Extractor)\n│   │   ├─ [任意] API01 JSON抽出 AUTH_TOKEN (JSON Extractor)\n│   │   ├─ [任意] API01 後処理 (JSR223 PostProcessor)\n│   │   ├─ [任意] API01 ステータス検証 200 (Response Assertion)\n│   │   ├─ [任意] API01 応答本文検証 (Response Assertion)\n│   │   └─ [任意] API01 応答時間検証 5000ms (Duration Assertion)\n│   ├─ API02_商品検索(GET) / API03_注文登録(POST form) / API04_クッキー確認(GET)\n│   └─ [任意] 変数確認 (Debug Sampler) ※既定は無効\n├─ [任意] 事後処理 (tearDown Thread Group)\n│   └─ [任意] テスト終了処理 (JSR223 Sampler)\n├─ [任意] 結果ツリー (View Results Tree)\n├─ [任意] 集計レポート (Summary Report)\n├─ [任意] 統計レポート (Aggregate Report)\n├─ [任意] JTL出力 全件 (Simple Data Writer)\n├─ [任意] エラー詳細出力 XML (Simple Data Writer)\n└─ [任意] 結果ログ出力 (JSR223 Listener)\n```\n\n### 要素の役割と削除時の影響\n<!-- cols: 30,28,52,44 -->\n| 要素 | JMeter の要素 | 役割 | 削除したときの影響 |\n|---|---|---|---|\n| [必須] 共通変数 | User Defined Variables | 接続先（PROTOCOL / HOST / PORT）と共通変数を定義。値は `${__P(名前,既定値)}` 形式で、実行時に -J で変更可能 | 削除不可（各 API が ${HOST} 等を参照） |\n| [任意] HTTPリクエスト初期値 | HTTP Request Defaults | 文字コードと接続 / 応答タイムアウト（CONNECT_TIMEOUT / RESPONSE_TIMEOUT） | タイムアウトが無制限になる |\n| [任意] 共通ヘッダ | HTTP Header Manager | 全 API に付与するヘッダ（設定 common_headers） | 共通ヘッダが送信されない |\n| [任意] クッキー管理 | HTTP Cookie Manager | 応答の Set-Cookie を保存し、以降のリクエストで送信。事前設定クッキーも送信 | クッキーが送信されない（セッションを使う API は失敗） |\n| [任意] 事前処理 | setUp Thread Group | テスト開始前に 1 回だけ実行する処理の入れ物 | 配下の処理（Excel 読込など）も削除される |\n| [任意] テスト開始処理 | JSR223 Sampler | 実行条件をログに出力し、集計値をリセット | 開始ログが出ない（動作には影響なし） |\n| [任意] Excelデータ読込 | JSR223 Sampler | Excel（.xlsx）をメモリに読み込む（Apache POI 不要） | Excel の値が変数に設定されない |\n| [必須] メイン処理 | Thread Group | API を実行する。THREADS / RAMP_UP / LOOPS（/ DURATION）で制御 | 削除不可 |\n| [任意] 連番 | Counter | 反復ごとに連番を ${SEQ} に設定（全スレッドで一意） | ${SEQ} が未定義になる |\n| [任意] 乱数 | Random Variable | 反復ごとに乱数を ${RANDOM_NO} に設定 | ${RANDOM_NO} が未定義になる |\n| [任意] CSVデータ読込 | CSV Data Set Config | CSV を 1 反復につき 1 行読み込み、列を変数に設定 | CSV の変数が未定義になる |\n| [任意] Excelデータ行取得 | JSR223 PreProcessor | Excel データから反復ごとに 1 行を取り出して変数に設定 | Excel の変数が未定義になる |\n| [任意] 共通前処理 | JSR223 PreProcessor | 全 API の直前に実行。${REQUEST_ID} ${TIMESTAMP} ${EPOCH_MS} を生成 | 左記の変数が未定義になる |\n| [任意] 共通後処理 | JSR223 PostProcessor | 全 API の直後に実行。${LAST_STATUS} ${LAST_ELAPSED_MS} を設定 | 左記の変数が未定義になる |\n| [任意] スループット制御 | Constant Throughput Timer | 1 分あたりの目標リクエスト数（THROUGHPUT）に調整。既定は無効 | なし |\n| APIxx_名前 | HTTP Request | API 定義 1 行分のリクエスト | その API が実行されない（配下の要素も削除される） |\n| [任意] APIxx ヘッダ | HTTP Header Manager | その API だけに付けるヘッダ | そのヘッダが送信されない |\n| [任意] APIxx 待機 | Constant Timer | API 呼び出し前の待機時間 | 待機しなくなる |\n| [任意] APIxx 前処理 / 後処理 | JSR223 PreProcessor / PostProcessor | API ごとの前処理・後処理（Groovy） | その処理が行われない |\n| [任意] APIxx 正規表現抽出 / JSON抽出 | Regular Expression Extractor / JSON Extractor | 応答から値を取り出して変数に設定 | 後続の API でその変数が未定義になる |\n| [任意] APIxx ステータス検証 / 応答本文検証 / 応答時間検証 | Response Assertion / Duration Assertion | 応答の検証。失敗はエラーとして集計 | 検証されなくなる |\n| [任意] 変数確認 | Debug Sampler | 変数の一覧を結果ツリーに表示（デバッグ用）。既定は無効 | なし |\n| [任意] 事後処理 / テスト終了処理 | tearDown Thread Group / JSR223 Sampler | 合計・成功・失敗件数をログに出力し、結果ログを閉じる | 終了時のサマリが出ない |\n| [任意] 結果ツリー / 集計レポート / 統計レポート | View Results Tree / Summary Report / Aggregate Report | GUI での結果確認用（ファイル出力なし） | GUI で結果が見られない（CLI 実行には影響なし） |\n| [任意] JTL出力 全件 | Simple Data Writer | 全リクエストの結果を result_all.jtl（CSV）に出力 | JTL が出力されない（run コマンドは別途 result.jtl を出力） |\n| [任意] エラー詳細出力 XML | Simple Data Writer | 失敗したリクエストのみ、リクエスト・レスポンス全文を error_detail.xml に出力 | エラー詳細が出力されない |\n| [任意] 結果ログ出力 | JSR223 Listener | 1 リクエスト 1 行の結果ログ api_result.log を出力 | 結果ログが出力されない |\n\n### 記入例で使っている変数の流れ\n<!-- cols: 30,34,54 -->\n| 変数 | 設定する要素 | 使っている場所（記入例） |\n|---|---|---|\n| PROTOCOL / HOST / PORT | [必須] 共通変数 | 全 API の接続先 |\n| SEQ | [任意] 連番 | API01 のボディ、API03 の orderNo、API04 のクエリ |\n| RANDOM_NO | [任意] 乱数 | API02 のクエリ rand |\n| USER_ID / PASSWORD / KEYWORD | [任意] CSVデータ読込 | API01 のボディ、API02 のクエリ |\n| ITEM_CODE / ITEM_NAME / QUANTITY | [任意] Excelデータ行取得 | API02 のクエリ、API03 のボディ |\n| REQUEST_ID / TIMESTAMP | [任意] 共通前処理 | API01 のヘッダ・ボディ |\n| SIGNATURE | API01 前処理 | API01 のヘッダ X-Signature |\n| AUTH_TOKEN | API01 JSON抽出 | API02 / API03 の Authorization ヘッダ |\n| ORDER_NO | API03 正規表現抽出 | API03 後処理のログ出力 |\n| CLIENT_ID | 共通変数（設定 user_variables） | 共通ヘッダ X-Client-Id |\n\n### 要素を削除・無効化するときのポイント\n- 一時的に外す場合は、削除より「Disable」（右クリック → Disable、または Ctrl + T）が安全です。\n- 変数を設定している要素を削除すると、その変数を使っている箇所では ${変数名} という文字列がそのまま送信されます。上の「変数の流れ」を確認してから削除してください。\n- CSV データ読込は、ファイルが見つからないとスレッドがエラーで停止します。CSV を使わない場合は削除してください。\n- 生成時に設定を OFF にすると、その要素は最初から生成されません（「4.入力項目_設定」参照）。\n\n## 7.実行方法\n<!-- summary: GUI / CLI での実行方法、run コマンド、実行時に変更できるプロパティ -->\n\n### 実行モードの使い分け\n<!-- cols: 16,46,56 -->\n| モード | 用途 | コマンド例 |\n|---|---|---|\n| GUI モード | シナリオの作成・動作確認・デバッグ（少量の実行） | `jmeter -t output/api_test.jmx` |\n| CLI モード | 負荷試験・大量件数の実行・定期実行 | `jmeter -n -t output/api_test.jmx -l result.jtl -e -o report` |\n\n### run コマンド（推奨）\nrun コマンドは CLI モードで JMeter を実行し、結果フォルダ（result/JMX名_日時/）に JTL・ログ・HTML レポートをまとめて出力し、最後に集計結果を表示します。\n\n```bash\n# RHEL 9.8\n# 既定（1スレッド × 1回）\n./jmeter_api_tool.sh run -t output/api_test.jmx\n# 10 スレッド × 100 回（1 API あたり 1,000 件）、ランプアップ 10 秒\n./jmeter_api_tool.sh run -t output/api_test.jmx --threads 10 --loops 100 --rampup 10\n# 接続先を切り替え、ヒープを 2GB にして実行\n./jmeter_api_tool.sh run -t output/api_test.jmx --threads 50 --loops 200 -J HOST=stg.example.com --heap 2g\n```\n\n```powershell\n# Windows 11\n.\\jmeter_api_tool.bat run -t output\\api_test.jmx\n.\\jmeter_api_tool.bat run -t output\\api_test.jmx -Threads 10 -Loops 100 -RampUp 10\n.\\jmeter_api_tool.bat run -t output\\api_test.jmx -Threads 50 -Loops 200 -Prop HOST=stg.example.com -Heap 2g\n```\n\n<!-- cols: 24,30,64 -->\n| sh のオプション | PowerShell のオプション | 説明 |\n|---|---|---|\n| -t ファイル | -Jmx（-t） | 実行する JMX。必須 |\n| --threads 数 | -Threads | スレッド数（THREADS） |\n| --rampup 秒 | -RampUp | ランプアップ秒数（RAMP_UP） |\n| --loops 数 | -Loops | ループ回数（LOOPS）。-1 で無限 |\n| --duration 秒 | -Duration | 実行時間（DURATION）。設定 duration を指定して生成した JMX のみ有効 |\n| -J 名前=値 | -Prop 名前=値,名前=値 | 任意のプロパティ。sh は複数回指定可、PowerShell はカンマ区切り |\n| -o フォルダ | -OutDir（-o） | 結果の出力先（既定: result） |\n| --heap サイズ | -Heap | JVM のヒープサイズ（例: 2g）。既定 1g |\n| --jmeter-home フォルダ | -JMeterHome | JMeter のフォルダ。省略時は環境変数 JMETER_HOME → PATH → 標準的な場所の順に探索 |\n| --fail-on-error | -FailOnError | 1 件でもエラーがあれば終了コード 2 を返す（CI 用） |\n| -- 以降の引数 | -ExtraArgs \"引数\" | JMeter にそのまま渡す追加オプション（例: プロキシ -H proxy.example.com -P 8080） |\n\n### 実行時に変更できるプロパティ\n<!-- cols: 34,32,52 -->\n| プロパティ | 既定値 | 説明 |\n|---|---|---|\n| THREADS | 設定 threads | 同時に実行するスレッド数（仮想ユーザー数） |\n| RAMP_UP | 設定 ramp_up | 全スレッドが起動するまでの秒数 |\n| LOOPS | 設定 loops | 各スレッドの繰り返し回数（-1 = 無限） |\n| DURATION | 設定 duration | 実行時間（秒）。duration を指定して生成した JMX のみ |\n| THROUGHPUT | 設定 throughput_per_min | スループット制御を有効にした場合の目標（件 / 分） |\n| PROTOCOL / HOST / PORT | URL から自動設定 | 接続先。2 つ目以降のホストは HOST_2、HOST_3 … |\n| CONNECT_TIMEOUT / RESPONSE_TIMEOUT | 10000 / 30000 | 接続 / 応答タイムアウト（ミリ秒） |\n| CSV_FILE | data/testdata.csv | CSV データファイル（相対パスは JMX のフォルダ基準） |\n| EXCEL_FILE / EXCEL_SHEET | data/testdata.xlsx / 先頭シート | Excel データファイルとシート名 |\n| RESULT_DIR | ~/result | JTL・結果ログの出力先（~/ は JMX のフォルダ。run コマンドは自動設定） |\n| LOG_RESPONSE | error | 結果ログへの本文出力（none / error / all） |\n| LOG_MAX_CHARS | 1000 | 結果ログに出力する本文の最大文字数 |\n| 共通変数の名前（例: ENV） | 設定 user_variables | 設定で定義した共通変数 |\n\n### JMeter を直接実行する場合\n```bash\n# RHEL 9.8（-f は既存の結果ファイル・レポートフォルダを削除してから実行）\njmeter -n -t output/api_test.jmx -l result/result.jtl -j result/jmeter.log \\\n  -e -o result/report -f -JTHREADS=10 -JLOOPS=100 -JRESULT_DIR=result\n```\n```powershell\n# Windows 11\n& \"$env:JMETER_HOME\\bin\\jmeter.bat\" -n -t output\\api_test.jmx -l result\\result.jtl -j result\\jmeter.log -e -o result\\report -f -JTHREADS=10 -JLOOPS=100 -JRESULT_DIR=result\n```\n- RESULT_DIR を相対パスで指定した場合は、JMeter を起動したフォルダ（カレントディレクトリ）が基準になります。\n- HTML レポートの出力先（-o）は、存在しないか空のフォルダである必要があります（-f を付けると自動で削除されます）。\n\n### GUI モードで実行する場合\n- メニューの「Run」→「Start」（Ctrl + R）で実行し、「結果ツリー」「集計レポート」で結果を確認します。\n- GUI で実行した場合も、JTL・結果ログは RESULT_DIR（既定: JMX と同じフォルダの result）に出力されます。\n- GUI での大量実行は避けてください。画面の描画で JMeter 自体が遅くなり、正しく測定できません。\n\n## 8.大量実行\n<!-- summary: 件数の考え方、負荷試験の推奨設定、JVM・OS の調整 -->\n\n### リクエスト件数の考え方\n総リクエスト件数 = THREADS（スレッド数）× LOOPS（ループ回数）× 有効な API の数\n\n<!-- cols: 22,50,46 -->\n| 目的 | 指定例 | 件数（API 4 本の場合） |\n|---|---|---|\n| 疎通確認（既定） | THREADS=1, LOOPS=1 | 4 件 |\n| 少量の連続実行 | THREADS=1, LOOPS=100 | 400 件 |\n| 並列実行 | THREADS=10, LOOPS=100 | 4,000 件 |\n| 時間指定の負荷試験 | 設定 duration=600 で生成し、THREADS=50, LOOPS=-1 | 600 秒間継続 |\n| CSV の行数だけ実行 | LOOPS=-1、設定 csv_recycle=false・csv_stop_thread=true | CSV 行数 × API 数 |\n\n### 負荷試験の推奨設定\n- CLI モード（run コマンド、または jmeter -n）で実行します。\n- [任意] 結果ログ出力は 1 件ごとにファイルへ書き込むため、毎秒数百件を超える場合は無効化するか、LOG_RESPONSE=none にします。\n- [任意] JTL出力 全件 は run コマンドの result.jtl と同じ内容を別ファイルに書き込むため、大量実行時は無効化してかまいません。\n- [任意] 結果ツリー / 集計レポート / 統計レポートは CLI では負荷になりにくいですが、不要なら削除または無効化します。\n- 待機時間（Constant Timer）で実際の操作間隔を再現し、スループット制御で目標の件数 / 分に合わせます。\n- ランプアップは「スレッド数 ÷ 1 秒あたりに増やしたいスレッド数」を目安にし、急激に負荷をかけないようにします。\n- Excel データは全行をメモリに読み込むため、数万行を超えるデータは CSV を使ってください。\n\n### JVM のヒープサイズ\n<!-- cols: 22,30,66 -->\n| 規模の目安 | ヒープ | 指定方法 |\n|---|---|---|\n| 〜100 スレッド | 1g（既定） | 指定不要 |\n| 〜500 スレッド | 2g 〜 4g | run コマンドの --heap 4g / -Heap 4g |\n| 500 スレッド超 | 4g 以上、または分散実行 | 同上。JMeter を直接実行する場合は環境変数 HEAP（例: -Xms4g -Xmx4g -XX:MaxMetaspaceSize=256m） |\n\n### OS の調整（RHEL 9.8）\n```bash\n# ファイルディスクリプタの上限（スレッド数が多い場合）\nulimit -n                 # 現在値の確認\nulimit -n 65535           # 一時的に変更（恒久的には /etc/security/limits.conf）\n# エフェメラルポートの範囲\nsysctl net.ipv4.ip_local_port_range\n```\n\n### OS の調整（Windows 11）\n- 動的ポートの範囲は `netsh int ipv4 show dynamicport tcp` で確認できます。\n- 長時間の試験では、スリープ・省電力の設定を無効にしてから実行します。\n- ウイルス対策ソフトのリアルタイムスキャンが結果ファイルの書き込みを遅くする場合があります（除外設定は社内のルールに従ってください）。\n\n### 分散実行（参考）\n1 台で必要な負荷をかけられない場合は、JMeter の分散実行（コントローラ 1 台 + 負荷生成サーバー複数台）を利用します。各サーバーで `jmeter-server` を起動し、`jmeter -n -t output/api_test.jmx -R サーバー1,サーバー2` で実行します。CSV / Excel のデータファイルは各サーバーの同じパスに配置が必要です。詳細は JMeter 公式マニュアルの「Remote Testing」を参照してください。\n\n## 9.結果確認\n<!-- summary: 出力ファイルの場所と見方（HTML レポート・JTL・ログ） -->\n\n### run コマンドの出力先\n```text\nresult/api_test_20261011_103000/\n├─ result.jtl          全リクエストの結果（CSV 形式）\n├─ jmeter.log          JMeter の実行ログ\n├─ report/index.html   HTML ダッシュボードレポート\n├─ api_result.log      結果ログ（[任意] 結果ログ出力）\n├─ result_all.jtl      JTL（[任意] JTL出力 全件）\n└─ error_detail.xml    失敗したリクエストの詳細（[任意] エラー詳細出力。失敗がある場合のみ）\n```\n\n### HTML レポート（report/index.html）\n<!-- cols: 28,90 -->\n| 画面 | 確認するポイント |\n|---|---|\n| Dashboard | APDEX（応答の満足度）、成功 / 失敗の割合、Statistics（API ごとの件数・エラー率・平均・90/95/99 パーセンタイル・スループット）、Errors（エラーの種類） |\n| Charts → Over Time | 応答時間・スループットの時間推移（性能劣化やエラー急増のタイミング） |\n| Charts → Throughput | 1 秒あたりのリクエスト数、応答コード別の件数 |\n| Charts → Response Times | 応答時間の分布、パーセンタイル |\n\nJTL から HTML レポートを作り直す場合は `jmeter -g result.jtl -o report2` を実行します。\n\n### JTL（result.jtl）の主な列\n<!-- cols: 22,96 -->\n| 列 | 内容 |\n|---|---|\n| timeStamp | 開始時刻（エポックミリ秒） |\n| elapsed | 応答時間（ミリ秒） |\n| label | API 名（HTTP リクエストの名前） |\n| responseCode / responseMessage | HTTP ステータスとメッセージ |\n| threadName | スレッド名 |\n| success | 成否（true / false） |\n| failureMessage | 検証（Assertion）エラーの内容 |\n| bytes / sentBytes | 受信 / 送信バイト数 |\n| grpThreads / allThreads | 実行中のスレッド数 |\n| URL | リクエストの URL |\n| Latency / Connect / IdleTime | 最初の応答までの時間 / 接続時間 / 待機時間（ミリ秒） |\n\n### 結果ログ（api_result.log）\nタブ区切りのテキストで、1 リクエスト 1 行です。列は「日時・結果（OK / NG）・ステータス・応答時間ms・受信バイト・ラベル・リクエスト（メソッドと URL）・スレッド・メッセージ」です。失敗時（LOG_RESPONSE=all の場合は全件）は次の行に [REQUEST] / [RESPONSE] として送信内容と応答本文（LOG_MAX_CHARS 文字まで）を出力します。最終行には合計・成功・失敗件数が出力されます。\n\n### エラー詳細（error_detail.xml）\n失敗したリクエストだけを、リクエストヘッダ・本文とレスポンスヘッダ・本文付きで出力します。JMeter GUI の「結果ツリー（View Results Tree）」の「Filename」欄でこのファイルを開くと、内容を画面で確認できます。\n\n### jmeter.log\n- 「==== テスト開始」「==== テスト終了」の行に、実行条件と合計・成功・失敗件数が出力されます。\n- 失敗したリクエストは「[API NG]」で始まる警告として出力されます。\n- スクリプトのエラーは ERROR / WARN として出力されます。うまく動かない場合はまずこのファイルを確認してください。\n\n## 10.変数と関数\n<!-- summary: 生成される変数、よく使う関数、JSR223（Groovy）・抽出の書き方 -->\n\n### 生成される変数\n<!-- cols: 34,26,58 -->\n| 変数 | 設定する要素 | 内容 |\n|---|---|---|\n| `${PROTOCOL}` `${HOST}` `${PORT}` | 共通変数 | 接続先 |\n| `${SEQ}` | 連番 | 000001 から始まる連番（全スレッドで一意） |\n| `${RANDOM_NO}` | 乱数 | 1 〜 99999 の乱数 |\n| `${REQUEST_ID}` | 共通前処理 | リクエストごとの UUID |\n| `${TIMESTAMP}` / `${EPOCH_MS}` | 共通前処理 | 現在時刻（ISO 8601 形式 / エポックミリ秒） |\n| `${LAST_STATUS}` / `${LAST_ELAPSED_MS}` | 共通後処理 | 直前の API のステータスと応答時間 |\n| `${EXCEL_ROW_NO}` | Excelデータ行取得 | 使用中の Excel の行番号 |\n| CSV / Excel の列名 | CSVデータ読込 / Excelデータ行取得 | 1 行目の列名がそのまま変数名になる |\n\n### よく使う関数\n<!-- cols: 18,46,54 -->\n| 関数 | 記入例 | 結果 |\n|---|---|---|\n| `__P` | `${__P(THREADS,1)}` | プロパティの値（未指定なら既定値） |\n| `__UUID` | `${__UUID()}` | ランダムな UUID |\n| `__time` | `${__time(yyyyMMddHHmmss,)}` | 現在時刻（書式指定） |\n| `__Random` | `${__Random(1,100,)}` | 1 〜 100 の乱数 |\n| `__RandomString` | `${__RandomString(8,abcdef0123456789,)}` | 指定した文字からなるランダムな文字列 |\n| `__counter` | `${__counter(FALSE,)}` | 全スレッド共通の連番（TRUE にするとスレッドごと） |\n| `__threadNum` | `${__threadNum}` | スレッド番号 |\n| `__urlencode` | `${__urlencode(${KEYWORD})}` | URL エンコード |\n| `__digest` | `${__digest(SHA-256,${USER_ID},,,)}` | ハッシュ値 |\n| `__timeShift` | `${__timeShift(yyyy-MM-dd,,P1D,,)}` | 日付の計算（例: 翌日） |\n| `__groovy` | `${__groovy(vars.get('SEQ').toInteger() * 2,)}` | Groovy の式の結果 |\n| `__V` | `${__V(USER_${SEQ})}` | 変数名を組み立てて参照 |\n\n### JSR223（Groovy）の書き方\n```groovy\n// 変数の取得・設定（値は文字列）\nString userId = vars.get('USER_ID')\nvars.put('USER_ID_UPPER', userId.toUpperCase())\n\n// プロパティ（全スレッド共通）の取得・設定\nString host = props.getProperty('HOST', 'localhost')\nprops.put('SHARED_TOKEN', vars.get('AUTH_TOKEN'))\n\n// 直前の応答（後処理・リスナーで使用）\nString code = prev.getResponseCode()\ndef json = new groovy.json.JsonSlurper().parseText(prev.getResponseDataAsString())\nvars.put('ORDER_ID', String.valueOf(json.data.id))\n\n// ログ出力（jmeter.log）\nlog.info('order id = ' + vars.get('ORDER_ID'))\n```\n- JSR223 のスクリプト内では `${変数名}` を使わず、`vars.get('変数名')` を使います（${} を使うとスクリプトのキャッシュが効かず、値によっては構文エラーになります）。\n- 言語は groovy、「Cache compiled script if available」はオンにします（生成した JMX は設定済み）。\n- vars はスレッドごと、props は全スレッド共通です。スレッドをまたいで値を共有する場合は props を使います。\n\n### 正規表現抽出の書き方\n<!-- cols: 40,78 -->\n| 抽出したいもの | 正規表現の例 |\n|---|---|\n| JSON の \"token\":\"abc\" の abc | `\"token\"\\s*:\\s*\"([^\"]+)\"` |\n| 数値の項目 \"id\": 123 の 123 | `\"id\"\\s*:\\s*(\\d+)` |\n| HTML の hidden 項目の値 | `name=\"csrf\" value=\"([^\"]+)\"` |\n| 応答ヘッダの値（GUI で対象を Response Headers に変更） | `Location: (.+)` |\n- ( ) で囲んだ部分が変数に入ります（テンプレート `$1$`）。見つからない場合は NOT_FOUND が入ります。\n- 一致したものをすべて取得するには、GUI で「Match No.」を -1 にします。`変数名_1`、`変数名_2` … と、件数を表す `変数名_matchNr` が作られます。\n\n### JSON 抽出（JSONPath）の書き方\n<!-- cols: 40,78 -->\n| 抽出したいもの | JSONPath |\n|---|---|\n| ルートの id | `$.id` |\n| 入れ子の値 | `$.data.user.name` |\n| 配列の先頭要素 | `$.items[0].id` |\n| 条件に一致する要素 | `$.items[?(@.status == 'active')].id` |\n| すべての id（Match No. を -1 にする） | `$..id` |\n\n## 11.トラブル対応\n<!-- summary: よくあるエラーと対処方法 -->\n<!-- cols: 36,36,48 -->\n| 症状 | 主な原因 | 対処 |\n|---|---|---|\n| `python3: command not found`（RHEL） | python3 が未導入 | `sudo dnf install -y python3` |\n| `$'\\r': command not found`（RHEL） | 改行コードが CRLF になっている | `sed -i 's/\\r$//' jmeter_api_tool.sh` |\n| このシステムではスクリプトの実行が無効になっている（Windows） | PowerShell の実行ポリシー | jmeter_api_tool.bat から実行する |\n| 日本語が文字化けする（コンソール） | 端末の文字コード | RHEL は `export LANG=ja_JP.UTF-8` など UTF-8 の端末を使用 |\n| JMeter が見つからない | JMETER_HOME が未設定 | 環境変数 JMETER_HOME を設定するか、--jmeter-home / -JMeterHome で指定 |\n| Java が見つからない・バージョンが古い | Java が未導入、PATH が未設定 | Java 17 をインストールし、JAVA_HOME と PATH を設定 |\n| `${USER_ID}` などがそのまま送信される | 変数を設定する要素が削除・無効、または変数名の誤り | 結果ツリーでリクエストを確認し、変数名と設定元の要素を見直す |\n| CSV の値が `<EOF>` になる | CSV の最終行に達した（csv_recycle=false） | csv_recycle=true にするか、ループ回数を見直す |\n| `File … must exist and be readable` | CSV ファイルが見つからない | CSV_FILE を JMX からの相対パス、または絶対パスで指定 |\n| 応答の日本語が文字化けする | 応答に文字コードの指定がない | JMeter 5.6.3 は既定で UTF-8。異なる場合は `-Jsampleresult.default.encoding=文字コード` |\n| Connection refused / Connect timed out | 接続先・ポートの誤り、ファイアウォール、プロキシ | URL を確認。プロキシ経由の場合は -H プロキシ -P ポート を追加（run は -- / -ExtraArgs で指定） |\n| SSLHandshakeException | TLS のバージョン・証明書の問題 | jmeter.log で詳細を確認。クライアント証明書が必要な場合は system.properties で keystore を設定 |\n| `java.lang.OutOfMemoryError` | ヒープ不足 | --heap / -Heap で増やす。CLI で実行し、結果ツリーを無効化 |\n| `BindException: Address already in use` | クライアント側のポート枯渇 | KeepAlive を有効（既定）にする、スレッド数を減らす、OS のポート範囲を広げる |\n| `Too many open files`（RHEL） | ファイルディスクリプタの上限 | `ulimit -n` で上限を引き上げる |\n| HTML レポートが作成されない | JTL が空、出力先フォルダが空でない | jmeter.log を確認。`jmeter -g result.jtl -o 新しいフォルダ` で再作成 |\n| Excel のデータが読み込まれない | ファイルパス・シート名の誤り、.xls 形式 | jmeter.log の [Excel] の行を確認。.xlsx 形式で保存する（.xls は非対応） |\n\n## 12.コマンド一覧\n<!-- summary: ツールと JMeter のコマンドラインオプション -->\n\n### ツールのコマンド\n<!-- cols: 12,46,46,30 -->\n| コマンド | RHEL 9.8（sh） | Windows 11（bat / ps1） | 説明 |\n|---|---|---|---|\n| init | `./jmeter_api_tool.sh init [-o フォルダ] [-f]` | `jmeter_api_tool.bat init [-OutDir フォルダ] [-Force]` | 入力テンプレートを作成（既定: input） |\n| generate | `./jmeter_api_tool.sh generate -i ファイル [-s 設定CSV] [-o フォルダ] [-n ファイル名] [--split] [-f]` | `jmeter_api_tool.bat generate -i ファイル [-s 設定CSV] [-OutDir フォルダ] [-JmxName ファイル名] [-Split] [-Force]` | JMX を生成（既定: output） |\n| manual | `./jmeter_api_tool.sh manual [-o フォルダ]` | `jmeter_api_tool.bat manual [-OutDir フォルダ]` | 本書（xlsx / md）を出力（既定: docs） |\n| run | `./jmeter_api_tool.sh run -t JMX [オプション]` | `jmeter_api_tool.bat run -t JMX [オプション]` | CLI モードで実行（既定: result） |\n| help | `./jmeter_api_tool.sh help` | `jmeter_api_tool.bat help` | ヘルプを表示 |\n\n### JMeter の主なオプション\n<!-- cols: 24,94 -->\n| オプション | 説明 |\n|---|---|\n| -n | CLI（非 GUI）モードで実行 |\n| -t ファイル | 実行する JMX |\n| -l ファイル | 結果（JTL）の出力先 |\n| -j ファイル | ログ（jmeter.log）の出力先 |\n| -e | テスト終了後に HTML レポートを作成 |\n| -o フォルダ | HTML レポートの出力先（存在しないか空のフォルダ） |\n| -f | 既存の結果ファイル・レポートフォルダを削除してから実行 |\n| -g ファイル | 既存の JTL から HTML レポートだけを作成（-o と併用） |\n| -J名前=値 | JMeter プロパティを設定（${__P(名前,既定値)} で参照） |\n| -G名前=値 | 分散実行時に全サーバーへプロパティを設定 |\n| -D名前=値 | Java のシステムプロパティを設定 |\n| -q ファイル | 追加のプロパティファイルを読み込む |\n| -H ホスト / -P ポート | プロキシサーバーのホスト / ポート |\n| -R ホスト1,ホスト2 | 分散実行（指定したサーバーで実行） |\n| -L レベル | ログレベル（例: -LDEBUG） |\n| -v | バージョンを表示 |\n",
 "template_guide_md": "## 記入方法（API定義テンプレート）\n\n### 基本的な使い方\n1. 「API一覧」シートに、呼び出す API を 1行 = 1API で記入します。上から順番に実行されます。\n2. 「設定」シートの「設定値」列で、スレッド数・ループ回数などの全体設定を変更します（空欄は既定値）。\n3. 保存後、次のコマンドで JMX を生成します。\n\n```text\nRHEL 9.8  : ./jmeter_api_tool.sh generate -i input/api_definition.xlsx\nWindows 11: jmeter_api_tool.bat generate -i input\\api_definition.xlsx\n```\n\n### 記入のルール\n- 「API一覧」の 4 行目（No 列が # で始まる行）は記入説明のため読み込まれません。API は 5 行目以降に記入します。行の追加・削除は自由です。\n- No 列が # で始まる行はコメントとして読み飛ばします。一時的に外したい API は No の先頭に # を付けるか、「有効」を N にします。\n- 複数の値を書く列（クエリパラメータ、ヘッダ、正規表現抽出など）は、セル内改行（Alt + Enter）で区切ります。\n- URL・パラメータ・ヘッダ・ボディには `${変数名}` で変数を埋め込めます。\n- シート名（API一覧・設定）と見出し行の列名は変更しないでください。列の並び順は入れ替えても読み込めます。\n- 入力セルは書式を「文字列」にしてあるため、0 始まりの値やカンマを含む値もそのまま保存されます。\n\n### API一覧の列\n<!-- @API_COLUMNS_TABLE@ -->\n\n### 記入例で使っている変数\n<!-- cols: 30,34,56 -->\n| 変数 | 設定する要素 | 内容 |\n|---|---|---|\n| `${USER_ID}` `${PASSWORD}` `${KEYWORD}` | CSVデータ読込（data/testdata.csv） | CSV の 1 行目の列名がそのまま変数名 |\n| `${ITEM_CODE}` `${ITEM_NAME}` `${QUANTITY}` | Excelデータ行取得（data/testdata.xlsx） | Excel の 1 行目の列名がそのまま変数名 |\n| `${SEQ}` / `${RANDOM_NO}` | 連番 / 乱数 | 反復ごとの連番（000001〜）/ 乱数 |\n| `${REQUEST_ID}` `${TIMESTAMP}` | 共通前処理 | リクエストごとの UUID / 現在時刻 |\n| `${AUTH_TOKEN}` | API01 の JSON抽出 | 前の API の応答から取り出した値 |\n| `${CLIENT_ID}` | 設定 user_variables | 共通変数 |\n\n### 生成されるもの\n- 生成された JMX の要素名は [必須] / [任意] で始まります。[任意] の要素は不要なら JMeter GUI で削除できます。\n- 詳しい使い方は利用方法（docs/JMeter_Usage_Guide.xlsx / .md）を参照してください。\n"
}
'@
$script:Spec = $script:SpecJson | ConvertFrom-Json
$script:Tool = $script:Spec.tool
$script:Xl = $script:Spec.xlsx
$script:Tpl = $script:Spec.template
$script:U = $script:Xl.units
$script:UUpper = [int]$script:U.upper; $script:UDigit = [int]$script:U.digit; $script:ULower = [int]$script:U.lower
$script:UPunct = [int]$script:U.punct; $script:ULatin = [int]$script:U.latin; $script:UHankaku = [int]$script:U.hankaku
$script:UKana = [int]$script:U.kana; $script:UWide = [int]$script:U.wide
$script:NsMain = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:NsRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:Ord = [System.StringComparison]::Ordinal
$script:Inv = [System.Globalization.CultureInfo]::InvariantCulture

# ----------------------------------------------------------------------------- console / errors
function Write-Info([string]$m) { Write-Host ('[INFO] ' + $m) }
function Write-Warn([string]$m) {
    $script:Warnings.Add($m)
    Write-Host ('[WARN] ' + $m) -ForegroundColor Yellow
}
function New-ToolError([string]$m) { return (New-Object System.InvalidOperationException $m) }

function Has-Prop($obj, [string]$name) {
    return ($null -ne $obj) -and ($null -ne $obj.PSObject.Properties[$name])
}
function Get-Prop($obj, [string]$name, $default) {
    if (Has-Prop $obj $name) { return $obj.PSObject.Properties[$name].Value }
    return $default
}

# ----------------------------------------------------------------------------- text utilities
$script:BadXml = New-Object System.Text.RegularExpressions.Regex '[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]'

function NormNl([string]$s) { return $s.Replace("`r`n", "`n").Replace("`r", "`n") }
function Clean([string]$s) { return $script:BadXml.Replace((NormNl $s), '') }
function XEsc([string]$s) {
    return (Clean $s).Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&apos;')
}
function Fill([string]$t, $values) {
    $s = $t
    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($k in $values.Keys) { $keys.Add([string]$k) }
    $keys.Sort([System.StringComparer]::Ordinal)
    foreach ($k in $keys) { $s = $s.Replace('@@' + $k + '@@', [string]$values[$k]) }
    return $s
}
function NormalizeKey([string]$s) {
    $n = $s.Normalize([System.Text.NormalizationForm]::FormKC)
    return ([regex]::Replace($n, '\s+', '')).ToLowerInvariant()
}
function JavaHash([string]$s) {
    [long]$h = 0
    foreach ($ch in $s.ToCharArray()) { $h = (31 * $h + [int]$ch) % 4294967296 }
    if ($h -ge 2147483648) { $h -= 4294967296 }
    return $h
}
function TextUnits([string]$s) {
    $u = 0
    $i = 0
    $n = $s.Length
    while ($i -lt $n) {
        $c = [int]$s[$i]
        if ($c -ge 0xD800 -and $c -le 0xDBFF -and ($i + 1) -lt $n) { $u += $script:UWide; $i += 2; continue }
        if ($c -lt 0x80) {
            if ($c -ge 0x41 -and $c -le 0x5A) { $u += $script:UUpper }
            elseif ($c -ge 0x30 -and $c -le 0x39) { $u += $script:UDigit }
            elseif ($c -ge 0x61 -and $c -le 0x7A) { $u += $script:ULower }
            else { $u += $script:UPunct }
        }
        elseif ($c -lt 0x1100) { $u += $script:ULatin }
        elseif ($c -ge 0xFF61 -and $c -le 0xFF9F) { $u += $script:UHankaku }
        elseif ($c -ge 0x3040 -and $c -le 0x30FF) { $u += $script:UKana }
        else { $u += $script:UWide }
        $i++
    }
    return $u
}
function WrapLines([string]$text, [int]$width) {
    $usable = [long]([Math]::Max($script:UDigit, $width * 20 - 30)) * 100
    [long]$n = 0
    foreach ($line in (NormNl $text).Split([char]10)) {
        [long]$u = [long](TextUnits $line) * 105
        $k = [long][Math]::Floor(($u + $usable - 1) / $usable)
        if ($k -lt 1) { $k = 1 }
        $n += $k
    }
    return [int]$n
}
function RowH10([int]$lines, $pad = $null) {
    if ($null -eq $pad) { $pad = [int]$script:Xl.pad_h10 }
    return [int][Math]::Min($lines * [int]$script:Xl.line_h10 + [int]$pad, [int]$script:Xl.max_h10)
}
function Get-Today { return (Get-Date).ToString('yyyy-MM-dd', $script:Inv) }
function Get-Stamp { return (Get-Date).ToString('yyyyMMdd_HHmmss', $script:Inv) }

# ----------------------------------------------------------------------------- files / csv
function FullPath([string]$p) { return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p) }

function Read-TextAuto([string]$path) {
    $bytes = [System.IO.File]::ReadAllBytes((FullPath $path))
    $n = $bytes.Length
    if ($n -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return @((NormNl ([System.Text.Encoding]::UTF8.GetString($bytes, 3, $n - 3))), 'UTF-8(BOM)')
    }
    if ($n -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return @((NormNl ([System.Text.Encoding]::Unicode.GetString($bytes, 2, $n - 2))), 'UTF-16LE')
    }
    if ($n -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        return @((NormNl ([System.Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $n - 2))), 'UTF-16BE')
    }
    try {
        $strict = New-Object System.Text.UTF8Encoding($false, $true)
        return @((NormNl $strict.GetString($bytes)), 'UTF-8')
    } catch { }
    try {
        $sjis = [System.Text.Encoding]::GetEncoding(932, [System.Text.EncoderFallback]::ExceptionFallback, [System.Text.DecoderFallback]::ExceptionFallback)
        return @((NormNl $sjis.GetString($bytes)), 'Shift_JIS(CP932)')
    } catch {
        throw (New-ToolError ('文字コードを判別できません（UTF-8 または Shift_JIS で保存してください）: ' + $path))
    }
}
function Write-BytesFile([string]$path, [byte[]]$data) {
    $full = FullPath $path
    $dir = [System.IO.Path]::GetDirectoryName($full)
    if ($dir -and -not [System.IO.Directory]::Exists($dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
    try {
        [System.IO.File]::WriteAllBytes($full, $data)
    } catch {
        throw (New-ToolError ('ファイルに書き込めません（Excel 等で開いている場合は閉じてください）: ' + $full))
    }
}
function Write-TextFile([string]$path, [string]$text, [bool]$bom = $false) {
    $enc = New-Object System.Text.UTF8Encoding($false)
    $body = $enc.GetBytes($text)
    if ($bom) {
        $ms = New-Object System.IO.MemoryStream
        $ms.Write([byte[]](0xEF, 0xBB, 0xBF), 0, 3)
        $ms.Write($body, 0, $body.Length)
        $body = $ms.ToArray()
    }
    Write-BytesFile $path $body
}
function Parse-Csv([string]$text, [string]$delim = ',') {
    $rows = New-Object System.Collections.Generic.List[object]
    $row = New-Object System.Collections.Generic.List[string]
    $field = New-Object System.Text.StringBuilder
    $d = [int][char]$delim
    $inq = $false
    $started = $false
    $i = 0
    $n = $text.Length
    while ($i -lt $n) {
        $c = [int]$text[$i]
        if ($inq) {
            if ($c -eq 34) {
                if ($i + 1 -lt $n -and [int]$text[$i + 1] -eq 34) { [void]$field.Append('"'); $i += 2; continue }
                $inq = $false
            } else {
                [void]$field.Append($text[$i])
            }
            $i++
            continue
        }
        if ($c -eq 34 -and $field.Length -eq 0) {
            $inq = $true; $started = $true
        } elseif ($c -eq $d) {
            $row.Add($field.ToString()); [void]$field.Clear(); $started = $true
        } elseif ($c -eq 10) {
            $row.Add($field.ToString()); $rows.Add($row.ToArray())
            $row = New-Object System.Collections.Generic.List[string]
            [void]$field.Clear(); $started = $false
        } else {
            [void]$field.Append($text[$i]); $started = $true
        }
        $i++
    }
    if ($started -or $field.Length -gt 0 -or $row.Count -gt 0) {
        $row.Add($field.ToString()); $rows.Add($row.ToArray())
    }
    return ,$rows
}
function CsvField([string]$s) {
    if ($s.Contains(',') -or $s.Contains('"') -or $s.Contains("`n") -or $s.Contains("`r") -or ($s -cne $s.Trim(' '))) {
        return '"' + $s.Replace('"', '""') + '"'
    }
    return $s
}
function To-Csv($rows) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($r in $rows) {
        $first = $true
        foreach ($f in $r) {
            if (-not $first) { [void]$sb.Append(',') }
            [void]$sb.Append((CsvField ([string]$f)))
            $first = $false
        }
        [void]$sb.Append("`r`n")
    }
    return $sb.ToString()
}

# ----------------------------------------------------------------------------- xlsx reader
function Get-Children($node, [string]$name) {
    $list = New-Object System.Collections.Generic.List[object]
    if ($null -eq $node) { return ,$list }
    foreach ($c in $node.ChildNodes) {
        if ($c.NodeType -eq [System.Xml.XmlNodeType]::Element -and $c.LocalName -ceq $name) { $list.Add($c) }
    }
    return ,$list
}
function Add-XlText($el, $sb) {
    foreach ($c in $el.ChildNodes) {
        if ($c.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        if ($c.LocalName -ceq 'rPh') { continue }
        if ($c.LocalName -ceq 't') { [void]$sb.Append($c.InnerText) } else { Add-XlText $c $sb }
    }
}
function Get-XlText($el) {
    $sb = New-Object System.Text.StringBuilder
    Add-XlText $el $sb
    return $sb.ToString()
}
function ColIndex([string]$ref) {
    $n = 0
    foreach ($ch in $ref.ToCharArray()) {
        $c = [int]$ch
        if ($c -ge 65 -and $c -le 90) { $n = $n * 26 + ($c - 64) }
        elseif ($c -ge 97 -and $c -le 122) { $n = $n * 26 + ($c - 96) }
        else { break }
    }
    return $n - 1
}
function Load-ZipXml($zip, [string]$name) {
    $entry = $zip.GetEntry($name)
    if ($null -eq $entry) { return $null }
    $stream = $entry.Open()
    try {
        $settings = New-Object System.Xml.XmlReaderSettings
        $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $reader = [System.Xml.XmlReader]::Create($stream, $settings)
        $doc = New-Object System.Xml.XmlDocument
        $doc.PreserveWhitespace = $true
        $doc.Load($reader)
        $reader.Close()
        return ,$doc.DocumentElement
    } finally {
        $stream.Dispose()
    }
}
function Read-XlRows($root, $shared) {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($sd in (Get-Children $root 'sheetData')) {
        $last = 0
        foreach ($row in (Get-Children $sd 'row')) {
            $r = $row.GetAttribute('r')
            if ([regex]::IsMatch($r, '^[0-9]+$')) { $rnum = [int]$r } else { $rnum = $last + 1 }
            $last = $rnum
            $cells = @{}
            $nextc = 0
            $maxc = -1
            foreach ($c in (Get-Children $row 'c')) {
                $ref = $c.GetAttribute('r')
                if ($ref) { $ci = ColIndex $ref } else { $ci = $nextc }
                if ($ci -lt 0) { $ci = $nextc }
                $nextc = $ci + 1
                $t = $c.GetAttribute('t')
                $val = ''
                if ($t -ceq 'inlineStr') {
                    foreach ($x in (Get-Children $c 'is')) { $val = Get-XlText $x }
                } else {
                    $v = ''
                    foreach ($x in (Get-Children $c 'v')) { $v = $x.InnerText }
                    if ($t -ceq 's') {
                        $idx = 0
                        if ([int]::TryParse($v.Trim(), [ref]$idx) -and $idx -ge 0 -and $idx -lt $shared.Count) { $val = $shared[$idx] } else { $val = '' }
                    } elseif ($t -ceq 'b') {
                        if ($v.Trim() -ceq '1') { $val = 'TRUE' } else { $val = 'FALSE' }
                    } else {
                        $val = $v
                    }
                }
                $cells[$ci] = $val
                if ($ci -gt $maxc) { $maxc = $ci }
            }
            $arr = New-Object string[] ($maxc + 1)
            for ($k = 0; $k -le $maxc; $k++) { if ($cells.ContainsKey($k)) { $arr[$k] = [string]$cells[$k] } else { $arr[$k] = '' } }
            $rows.Add(@($rnum, $arr))
        }
    }
    return ,$rows
}
function Read-Xlsx([string]$path) {
    $full = FullPath $path
    $fs = $null
    $zip = $null
    try {
        $fs = New-Object System.IO.FileStream($full, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
    } catch {
        if ($null -ne $fs) { $fs.Dispose() }
        throw (New-ToolError ('Excel ファイルを開けません（.xlsx 形式で保存してください）: ' + $path))
    }
    try {
        $shared = New-Object System.Collections.Generic.List[string]
        $sst = Load-ZipXml $zip 'xl/sharedStrings.xml'
        foreach ($si in (Get-Children $sst 'si')) { $shared.Add((Get-XlText $si)) }
        $wb = Load-ZipXml $zip 'xl/workbook.xml'
        $rels = Load-ZipXml $zip 'xl/_rels/workbook.xml.rels'
        if ($null -eq $wb -or $null -eq $rels) { throw (New-ToolError ('Excel ファイルの形式が正しくありません: ' + $path)) }
        $relmap = New-Object 'System.Collections.Generic.Dictionary[string,string]'
        foreach ($r in (Get-Children $rels 'Relationship')) { $relmap[$r.GetAttribute('Id')] = $r.GetAttribute('Target') }
        $sheets = New-Object System.Collections.Generic.List[object]
        foreach ($sheetsEl in (Get-Children $wb 'sheets')) {
            foreach ($el in (Get-Children $sheetsEl 'sheet')) {
                $rid = ''
                foreach ($a in $el.Attributes) { if ($a.LocalName -ceq 'id' -and $a.NamespaceURI) { $rid = $a.Value } }
                $target = ''
                if ($relmap.ContainsKey($rid)) { $target = $relmap[$rid] }
                if (-not $target) { continue }
                if ($target.StartsWith('/', $script:Ord)) { $part = $target.Substring(1) } else { $part = 'xl/' + $target }
                $root = Load-ZipXml $zip $part
                if ($null -ne $root) { $rows = Read-XlRows $root $shared } else { $rows = New-Object System.Collections.Generic.List[object] }
                $sheets.Add(@($el.GetAttribute('name'), $rows))
            }
        }
        return ,$sheets
    } finally {
        $zip.Dispose()
        $fs.Dispose()
    }
}

# ----------------------------------------------------------------------------- xlsx writer
$script:XfIndex = @{}
$script:XfFont = @{}
$xfi = 0
foreach ($x in $script:Spec.styles.xfs) { $script:XfIndex[[string]$x[0]] = $xfi; $script:XfFont[[string]$x[0]] = [int]$x[1]; $xfi++ }

function ColLetter([int]$c) {
    $s = ''
    while ($c -gt 0) {
        $rem = ($c - 1) % 26
        $c = [int][Math]::Floor(($c - 1) / 26)
        $s = [string][char](65 + $rem) + $s
    }
    return $s
}
function CRef([int]$r, [int]$c) { return (ColLetter $c) + [string]$r }
function FmtH10([int]$h) {
    $a = [int][Math]::Floor($h / 10)
    $b = $h % 10
    if ($b -eq 0) { return [string]$a }
    return [string]$a + '.' + [string]$b
}
function PlainOf($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [string]) { return $v }
    $sb = New-Object System.Text.StringBuilder
    foreach ($run in $v) { [void]$sb.Append([string]$run[0]) }
    return $sb.ToString()
}
function New-Sheet([string]$name) {
    return @{
        Name = $name
        Cells = New-Object 'System.Collections.Generic.SortedDictionary[long,object]'
        H10 = New-Object 'System.Collections.Generic.SortedDictionary[int,int]'
        Widths = New-Object 'System.Collections.Generic.SortedDictionary[int,int]'
        Merges = New-Object System.Collections.Generic.List[object]
        Validations = New-Object System.Collections.Generic.List[object]
        Links = New-Object System.Collections.Generic.List[object]
        FreezeRows = 0
        FreezeCols = 0
        Selected = $false
        Paper = 9
        FitWidth = 1
    }
}
function Put-Cell($sh, [int]$r, [int]$c, $value, [string]$style) {
    $sh.Cells[[long]$r * 20000 + $c] = @($value, $style)
}
function Put-Block($sh, [int]$r, [int]$c1, [int]$c2, $value, [string]$style) {
    Put-Cell $sh $r $c1 $value $style
    for ($c = $c1 + 1; $c -le $c2; $c++) { Put-Cell $sh $r $c $null $style }
    if ($c2 -gt $c1) { $sh.Merges.Add(@($r, $c1, $r, $c2)) }
}
function T-Xml([string]$s, [bool]$force = $false) {
    $s = Clean $s
    $sp = ''
    if ($force -or ($s.Length -gt 0 -and (" `t`n".IndexOf($s[0]) -ge 0 -or " `t`n".IndexOf($s[$s.Length - 1]) -ge 0))) { $sp = ' xml:space="preserve"' }
    return '<t' + $sp + '>' + (XEsc $s) + '</t>'
}
function RPr-Xml([int]$fontIdx, [bool]$bold) {
    $f = $script:Spec.styles.fonts[$fontIdx]
    $b = $bold -or [bool](Get-Prop $f 'b' $false)
    $bx = ''
    if ($b) { $bx = '<b/>' }
    return '<rPr>' + $bx + '<sz val="' + [string]$f.sz + '"/><color rgb="' + $f.color + '"/><rFont val="' + $script:Xl.font_name + '"/><family val="3"/><charset val="128"/></rPr>'
}
function Pane-Xml([int]$fr, [int]$fc) {
    if ($fr -eq 0 -and $fc -eq 0) { return '' }
    $tl = CRef ($fr + 1) ($fc + 1)
    if ($fr -gt 0 -and $fc -gt 0) {
        return '<pane xSplit="' + $fc + '" ySplit="' + $fr + '" topLeftCell="' + $tl + '" activePane="bottomRight" state="frozen"/>' +
            '<selection pane="topRight"/><selection pane="bottomLeft"/>' +
            '<selection pane="bottomRight" activeCell="' + $tl + '" sqref="' + $tl + '"/>'
    }
    if ($fr -gt 0) {
        return '<pane ySplit="' + $fr + '" topLeftCell="' + $tl + '" activePane="bottomLeft" state="frozen"/>' +
            '<selection pane="bottomLeft" activeCell="' + $tl + '" sqref="' + $tl + '"/>'
    }
    return '<pane xSplit="' + $fc + '" topLeftCell="' + $tl + '" activePane="topRight" state="frozen"/>' +
        '<selection pane="topRight" activeCell="' + $tl + '" sqref="' + $tl + '"/>'
}
function Sheet-Xml($sh, $sst) {
    $rowsSet = New-Object 'System.Collections.Generic.SortedSet[int]'
    $maxr = 1
    $maxc = 1
    foreach ($key in $sh.Cells.Keys) {
        $r = [int][Math]::Floor($key / 20000)
        $c = [int]($key % 20000)
        [void]$rowsSet.Add($r)
        if ($r -gt $maxr) { $maxr = $r }
        if ($c -gt $maxc) { $maxc = $c }
    }
    foreach ($r in $sh.H10.Keys) { [void]$rowsSet.Add($r) }
    $o = New-Object System.Text.StringBuilder
    [void]$o.Append("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n")
    [void]$o.Append('<worksheet xmlns="' + $script:NsMain + '" xmlns:r="' + $script:NsRel + '">')
    [void]$o.Append('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>')
    [void]$o.Append('<dimension ref="A1:' + (CRef $maxr $maxc) + '"/>')
    $sel = ''
    if ($sh.Selected) { $sel = ' tabSelected="1"' }
    [void]$o.Append('<sheetViews><sheetView showGridLines="0"' + $sel + ' workbookViewId="0">' + (Pane-Xml $sh.FreezeRows $sh.FreezeCols) + '</sheetView></sheetViews>')
    [void]$o.Append('<sheetFormatPr defaultRowHeight="14.4"/>')
    if ($sh.Widths.Count -gt 0) {
        [void]$o.Append('<cols>')
        foreach ($c in $sh.Widths.Keys) { [void]$o.Append('<col min="' + $c + '" max="' + $c + '" width="' + $sh.Widths[$c] + '" customWidth="1"/>') }
        [void]$o.Append('</cols>')
    }
    [void]$o.Append('<sheetData>')
    $byRow = @{}
    foreach ($key in $sh.Cells.Keys) {
        $r = [int][Math]::Floor($key / 20000)
        if (-not $byRow.ContainsKey($r)) { $byRow[$r] = New-Object System.Collections.Generic.List[long] }
        $byRow[$r].Add($key)
    }
    foreach ($r in $rowsSet) {
        $attrs = ' r="' + $r + '"'
        if ($sh.H10.ContainsKey($r)) { $attrs += ' ht="' + (FmtH10 $sh.H10[$r]) + '" customHeight="1"' }
        if (-not $byRow.ContainsKey($r)) { [void]$o.Append('<row' + $attrs + '/>'); continue }
        [void]$o.Append('<row' + $attrs + '>')
        foreach ($key in $byRow[$r]) {
            $c = [int]($key % 20000)
            $entry = $sh.Cells[$key]
            $value = $entry[0]
            $style = [string]$entry[1]
            $s = $script:XfIndex[$style]
            if ($null -eq $value -or (PlainOf $value) -ceq '') {
                [void]$o.Append('<c r="' + (CRef $r $c) + '" s="' + $s + '"/>')
            } else {
                $idx = Get-SstIndex $sst $value $script:XfFont[$style]
                [void]$o.Append('<c r="' + (CRef $r $c) + '" s="' + $s + '" t="s"><v>' + $idx + '</v></c>')
            }
        }
        [void]$o.Append('</row>')
    }
    [void]$o.Append('</sheetData>')
    if ($sh.Merges.Count -gt 0) {
        [void]$o.Append('<mergeCells count="' + $sh.Merges.Count + '">')
        foreach ($m in $sh.Merges) { [void]$o.Append('<mergeCell ref="' + (CRef $m[0] $m[1]) + ':' + (CRef $m[2] $m[3]) + '"/>') }
        [void]$o.Append('</mergeCells>')
    }
    if ($sh.Validations.Count -gt 0) {
        [void]$o.Append('<dataValidations count="' + $sh.Validations.Count + '">')
        foreach ($v in $sh.Validations) {
            [void]$o.Append('<dataValidation type="list" errorStyle="warning" allowBlank="1" showErrorMessage="1" sqref="' + $v[0] + '">' +
                '<formula1>' + (XEsc ('"' + $v[1] + '"')) + '</formula1></dataValidation>')
        }
        [void]$o.Append('</dataValidations>')
    }
    if ($sh.Links.Count -gt 0) {
        [void]$o.Append('<hyperlinks>')
        foreach ($l in $sh.Links) {
            [void]$o.Append('<hyperlink ref="' + $l[0] + '" location="' + (XEsc ("'" + ([string]$l[1]).Replace("'", "''") + "'!A1")) + '" display="' + (XEsc $l[1]) + '"/>')
        }
        [void]$o.Append('</hyperlinks>')
    }
    [void]$o.Append('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
    [void]$o.Append('<pageSetup paperSize="' + $sh.Paper + '" orientation="landscape" fitToWidth="' + $sh.FitWidth + '" fitToHeight="0"/>')
    [void]$o.Append('<headerFooter><oddFooter>&amp;C&amp;P / &amp;N</oddFooter></headerFooter>')
    [void]$o.Append('</worksheet>')
    return $o.ToString()
}
function New-Sst {
    return @{
        Map = New-Object 'System.Collections.Generic.Dictionary[string,int]'
        Items = New-Object System.Collections.Generic.List[object]
        Refs = 0
    }
}
function Get-SstIndex($sst, $value, [int]$font) {
    $sst.Refs++
    if ($value -is [string]) {
        $key = 'S' + [char]1 + $value
    } else {
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append('R' + [char]1 + [string]$font + [char]1)
        $first = $true
        foreach ($run in $value) {
            if (-not $first) { [void]$sb.Append([char]2) }
            if ($run[1]) { [void]$sb.Append('1') } else { [void]$sb.Append('0') }
            [void]$sb.Append([string]$run[0])
            $first = $false
        }
        $key = $sb.ToString()
    }
    if ($sst.Map.ContainsKey($key)) { return $sst.Map[$key] }
    $i = $sst.Items.Count
    $sst.Map[$key] = $i
    $sst.Items.Add(@($value, $font))
    return $i
}
function Styles-Xml {
    $st = $script:Spec.styles
    $fn = $script:Xl.font_name
    $o = New-Object System.Text.StringBuilder
    [void]$o.Append("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n<styleSheet xmlns=`"" + $script:NsMain + "`">")
    [void]$o.Append('<fonts count="' + @($st.fonts).Count + '">')
    foreach ($f in $st.fonts) {
        $b = ''; $it = ''; $u = ''
        if ([bool](Get-Prop $f 'b' $false)) { $b = '<b/>' }
        if ([bool](Get-Prop $f 'i' $false)) { $it = '<i/>' }
        if ([bool](Get-Prop $f 'u' $false)) { $u = '<u/>' }
        [void]$o.Append('<font>' + $b + $it + $u + '<sz val="' + [string]$f.sz + '"/><color rgb="' + $f.color + '"/><name val="' + $fn + '"/><family val="3"/><charset val="128"/></font>')
    }
    [void]$o.Append('</fonts>')
    [void]$o.Append('<fills count="' + @($st.fills).Count + '">')
    $i = 0
    foreach ($color in $st.fills) {
        if ($i -eq 0) { [void]$o.Append('<fill><patternFill patternType="none"/></fill>') }
        elseif ($i -eq 1) { [void]$o.Append('<fill><patternFill patternType="gray125"/></fill>') }
        else { [void]$o.Append('<fill><patternFill patternType="solid"><fgColor rgb="' + $color + '"/><bgColor indexed="64"/></patternFill></fill>') }
        $i++
    }
    [void]$o.Append('</fills>')
    [void]$o.Append('<borders count="' + @($st.borders).Count + '">')
    foreach ($b in $st.borders) {
        [void]$o.Append('<border>')
        foreach ($side in @('left', 'right', 'top', 'bottom')) {
            if (Has-Prop $b $side) {
                $spec = $b.PSObject.Properties[$side].Value
                [void]$o.Append('<' + $side + ' style="' + $spec[0] + '"><color rgb="' + $spec[1] + '"/></' + $side + '>')
            } else {
                [void]$o.Append('<' + $side + '/>')
            }
        }
        [void]$o.Append('<diagonal/></border>')
    }
    [void]$o.Append('</borders>')
    [void]$o.Append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
    [void]$o.Append('<cellXfs count="' + @($st.xfs).Count + '">')
    foreach ($x in $st.xfs) {
        $attrs = 'numFmtId="' + [int]$x[8] + '" fontId="' + [int]$x[1] + '" fillId="' + [int]$x[2] + '" borderId="' + [int]$x[3] + '" xfId="0"'
        if ([int]$x[8] -ne 0) { $attrs += ' applyNumberFormat="1"' }
        $attrs += ' applyFont="1" applyFill="1" applyBorder="1"'
        $a = New-Object System.Collections.Generic.List[string]
        if ([string]$x[4]) { $a.Add('horizontal="' + $x[4] + '"') }
        if ([string]$x[5]) { $a.Add('vertical="' + $x[5] + '"') }
        if ([bool]$x[6]) { $a.Add('wrapText="1"') }
        if ([int]$x[7] -ne 0) { $a.Add('indent="' + [int]$x[7] + '"') }
        if ($a.Count -gt 0) {
            [void]$o.Append('<xf ' + $attrs + ' applyAlignment="1"><alignment ' + ($a -join ' ') + '/></xf>')
        } else {
            [void]$o.Append('<xf ' + $attrs + '/>')
        }
    }
    [void]$o.Append('</cellXfs>')
    [void]$o.Append('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')
    [void]$o.Append('<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>')
    [void]$o.Append('</styleSheet>')
    return $o.ToString()
}
function Save-Book($sheets, [string]$path, [string]$title) {
    $sst = New-Sst
    $sheetParts = New-Object System.Collections.Generic.List[string]
    foreach ($sh in $sheets) { $sheetParts.Add((Sheet-Xml $sh $sst)) }
    $o = New-Object System.Text.StringBuilder
    [void]$o.Append("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n<sst xmlns=`"" + $script:NsMain + "`" count=`"" + $sst.Refs + "`" uniqueCount=`"" + $sst.Items.Count + "`">")
    foreach ($item in $sst.Items) {
        $value = $item[0]
        if ($value -is [string]) {
            [void]$o.Append('<si>' + (T-Xml $value) + '</si>')
        } else {
            [void]$o.Append('<si>')
            foreach ($run in $value) { [void]$o.Append('<r>' + (RPr-Xml $item[1] ([bool]$run[1])) + (T-Xml ([string]$run[0]) $true) + '</r>') }
            [void]$o.Append('</si>')
        }
    }
    [void]$o.Append('</sst>')
    $sstXml = $o.ToString()
    $active = 0
    $n = 0
    foreach ($sh in $sheets) { if ($sh.Selected) { $active = $n }; $n++ }
    $ct = New-Object System.Text.StringBuilder
    [void]$ct.Append("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n")
    [void]$ct.Append('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
    [void]$ct.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
    [void]$ct.Append('<Default Extension="xml" ContentType="application/xml"/>')
    [void]$ct.Append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
    for ($i = 1; $i -le $n; $i++) {
        [void]$ct.Append('<Override PartName="/xl/worksheets/sheet' + $i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    }
    [void]$ct.Append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    [void]$ct.Append('<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>')
    [void]$ct.Append('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>')
    [void]$ct.Append('<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>')
    [void]$ct.Append('</Types>')
    $rels = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n" +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
        '</Relationships>'
    $now = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ', $script:Inv)
    $core = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n" +
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        '<dc:title>' + (XEsc $title) + '</dc:title><dc:creator>' + (XEsc ($script:Tool.name + ' ' + $script:ToolVersion)) + '</dc:creator>' +
        '<dcterms:created xsi:type="dcterms:W3CDTF">' + $now + '</dcterms:created>' +
        '<dcterms:modified xsi:type="dcterms:W3CDTF">' + $now + '</dcterms:modified>' +
        '</cp:coreProperties>'
    $app = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n" +
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
        '<Application>Microsoft Excel</Application></Properties>'
    $wb = New-Object System.Text.StringBuilder
    [void]$wb.Append("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n")
    [void]$wb.Append('<workbook xmlns="' + $script:NsMain + '" xmlns:r="' + $script:NsRel + '">')
    [void]$wb.Append('<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15000" activeTab="' + $active + '"/></bookViews>')
    [void]$wb.Append('<sheets>')
    $i = 1
    foreach ($sh in $sheets) { [void]$wb.Append('<sheet name="' + (XEsc $sh.Name) + '" sheetId="' + $i + '" r:id="rId' + $i + '"/>'); $i++ }
    [void]$wb.Append('</sheets></workbook>')
    $wr = New-Object System.Text.StringBuilder
    [void]$wr.Append("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n")
    [void]$wr.Append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    for ($i = 1; $i -le $n; $i++) {
        [void]$wr.Append('<Relationship Id="rId' + $i + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + $i + '.xml"/>')
    }
    [void]$wr.Append('<Relationship Id="rId' + ($n + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    [void]$wr.Append('<Relationship Id="rId' + ($n + 2) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>')
    [void]$wr.Append('</Relationships>')
    $parts = New-Object System.Collections.Generic.List[object]
    $parts.Add(@('[Content_Types].xml', $ct.ToString()))
    $parts.Add(@('_rels/.rels', $rels))
    $parts.Add(@('docProps/app.xml', $app))
    $parts.Add(@('docProps/core.xml', $core))
    $parts.Add(@('xl/workbook.xml', $wb.ToString()))
    $parts.Add(@('xl/_rels/workbook.xml.rels', $wr.ToString()))
    $parts.Add(@('xl/styles.xml', (Styles-Xml)))
    $parts.Add(@('xl/sharedStrings.xml', $sstXml))
    for ($i = 0; $i -lt $sheetParts.Count; $i++) { $parts.Add(@(('xl/worksheets/sheet' + ($i + 1) + '.xml'), $sheetParts[$i])) }
    $ms = New-Object System.IO.MemoryStream
    $zip = New-Object System.IO.Compression.ZipArchive($ms, [System.IO.Compression.ZipArchiveMode]::Create, $true)
    $enc = New-Object System.Text.UTF8Encoding($false)
    foreach ($p in $parts) {
        $entry = $zip.CreateEntry([string]$p[0], [System.IO.Compression.CompressionLevel]::Optimal)
        $s = $entry.Open()
        $bytes = $enc.GetBytes([string]$p[1])
        $s.Write($bytes, 0, $bytes.Length)
        $s.Dispose()
    }
    $zip.Dispose()
    Write-BytesFile $path $ms.ToArray()
}

# ----------------------------------------------------------------------------- markdown -> sheets
$script:ListRe = '^(\s*)([-*]|[0-9]+\.)\s+(.*)$'

function Md-SplitRow([string]$line) {
    $s = $line.Trim()
    if ($s.StartsWith('|', $script:Ord)) { $s = $s.Substring(1) }
    if ($s.EndsWith('|', $script:Ord) -and -not $s.EndsWith('\|', $script:Ord)) { $s = $s.Substring(0, $s.Length - 1) }
    $cells = New-Object System.Collections.Generic.List[string]
    $cur = New-Object System.Text.StringBuilder
    $i = 0
    while ($i -lt $s.Length) {
        if ($s[$i] -ceq '\' -and $i + 1 -lt $s.Length -and $s[$i + 1] -ceq '|') { [void]$cur.Append('\|'); $i += 2; continue }
        if ($s[$i] -ceq '|') { $cells.Add($cur.ToString().Trim()); [void]$cur.Clear() } else { [void]$cur.Append($s[$i]) }
        $i++
    }
    $cells.Add($cur.ToString().Trim())
    return ,($cells.ToArray())
}
function Md-Fix([string]$t, [bool]$br) {
    if ($br) { $t = $t.Replace('<br>', "`n") }
    return $t.Replace('\|', '|')
}
function Md-Inline([string]$text) {
    $text = [regex]::Replace($text, '\[([^\]]+)\]\(([^)]+)\)', '$1 ($2)')
    $runs = New-Object System.Collections.Generic.List[object]
    $pos = 0
    foreach ($m in [regex]::Matches($text, '`([^`]*)`|\*\*(.+?)\*\*')) {
        if ($m.Index -gt $pos) { $runs.Add(@((Md-Fix $text.Substring($pos, $m.Index - $pos) $true), $false)) }
        if ($m.Groups[1].Success) { $runs.Add(@((Md-Fix $m.Groups[1].Value $false), $false)) }
        else { $runs.Add(@((Md-Fix $m.Groups[2].Value $true), $true)) }
        $pos = $m.Index + $m.Length
    }
    if ($pos -lt $text.Length) { $runs.Add(@((Md-Fix $text.Substring($pos) $true), $false)) }
    $merged = New-Object System.Collections.Generic.List[object]
    $anyBold = $false
    foreach ($r in $runs) {
        if ([string]$r[0] -ceq '') { continue }
        if ($merged.Count -gt 0 -and $merged[$merged.Count - 1][1] -eq $r[1]) {
            $merged[$merged.Count - 1][0] = [string]$merged[$merged.Count - 1][0] + [string]$r[0]
        } else {
            $merged.Add(@([string]$r[0], [bool]$r[1]))
        }
        if ($r[1]) { $anyBold = $true }
    }
    if (-not $anyBold) {
        $sb = New-Object System.Text.StringBuilder
        foreach ($r in $merged) { [void]$sb.Append([string]$r[0]) }
        return $sb.ToString()
    }
    return ,($merged.ToArray())
}
function New-Block([string]$t) { return @{ t = $t } }
function Md-Parse([string]$md) {
    $lines = (NormNl $md).Split([char]10)
    $doc = @{ title = ''; cover = (New-Object System.Collections.Generic.List[object]); sections = (New-Object System.Collections.Generic.List[object]) }
    $blocks = $doc.cover
    $cols = $null
    $i = 0
    $n = $lines.Length
    while ($i -lt $n) {
        $line = $lines[$i]
        $s = $line.Trim()
        if ($s.StartsWith('```', $script:Ord)) {
            $code = New-Object System.Collections.Generic.List[string]
            $i++
            while ($i -lt $n -and -not $lines[$i].Trim().StartsWith('```', $script:Ord)) { $code.Add($lines[$i]); $i++ }
            $i++
            $b = New-Block 'code'; $b.lines = $code.ToArray(); $blocks.Add($b)
            continue
        }
        if ($s.StartsWith('<!--', $script:Ord)) {
            $text = $s
            while (-not $text.Contains('-->') -and $i + 1 -lt $n) { $i++; $text += "`n" + $lines[$i] }
            $i++
            $m = [regex]::Match($text, '^<!--\s*cols:\s*([0-9,\s]+?)\s*-->$')
            if ($m.Success) {
                $cl = New-Object System.Collections.Generic.List[int]
                foreach ($x in [regex]::Split($m.Groups[1].Value.Trim(), '[,\s]+')) { if ($x) { $cl.Add([int]$x) } }
                $cols = $cl.ToArray()
            }
            $m = [regex]::Match($text, '^<!--\s*summary:\s*([^\n]*?)\s*-->$')
            if ($m.Success -and $doc.sections.Count -gt 0) { $doc.sections[$doc.sections.Count - 1].summary = $m.Groups[1].Value }
            continue
        }
        if ($s -ceq '') { $i++; continue }
        if ($line.StartsWith('# ', $script:Ord)) { $doc.title = $line.Substring(2).Trim(); $i++; continue }
        if ($line.StartsWith('## ', $script:Ord)) {
            $sec = @{ title = $line.Substring(3).Trim(); summary = ''; blocks = (New-Object System.Collections.Generic.List[object]) }
            $doc.sections.Add($sec)
            $blocks = $sec.blocks
            $i++
            continue
        }
        if ($line.StartsWith('### ', $script:Ord)) { $b = New-Block 'h3'; $b.text = $line.Substring(4).Trim(); $blocks.Add($b); $i++; continue }
        if ($line.StartsWith('#### ', $script:Ord)) { $b = New-Block 'h4'; $b.text = $line.Substring(5).Trim(); $blocks.Add($b); $i++; continue }
        if ($s.StartsWith('|', $script:Ord)) {
            $rows = New-Object System.Collections.Generic.List[object]
            while ($i -lt $n -and $lines[$i].Trim().StartsWith('|', $script:Ord)) { $rows.Add((Md-SplitRow $lines[$i])); $i++ }
            $sep = $false
            if ($rows.Count -gt 1) {
                $sep = $true
                foreach ($c in $rows[1]) { if ($c -cne '' -and -not [regex]::IsMatch($c, '^:?-+:?$')) { $sep = $false } }
            }
            $b = New-Block 'table'
            $b.header = $rows[0]
            $body = New-Object System.Collections.Generic.List[object]
            $startAt = 1
            if ($sep) { $startAt = 2 }
            for ($k = $startAt; $k -lt $rows.Count; $k++) { $body.Add($rows[$k]) }
            $b.rows = $body
            $b.cols = $cols
            $blocks.Add($b)
            $cols = $null
            continue
        }
        if ([regex]::IsMatch($line, $script:ListRe)) {
            $items = New-Object System.Collections.Generic.List[object]
            while ($i -lt $n) {
                $m = [regex]::Match($lines[$i], $script:ListRe)
                if (-not $m.Success) { break }
                $items.Add(@{ level = [int][Math]::Floor($m.Groups[1].Value.Length / 2); mark = $m.Groups[2].Value; text = $m.Groups[3].Value.Trim() })
                $i++
            }
            $b = New-Block 'list'; $b.items = $items; $blocks.Add($b)
            continue
        }
        if ($s.StartsWith('>', $script:Ord)) {
            $q = New-Object System.Collections.Generic.List[string]
            while ($i -lt $n -and $lines[$i].Trim().StartsWith('>', $script:Ord)) { $q.Add([regex]::Replace($lines[$i], '^\s*>\s?', '')); $i++ }
            $b = New-Block 'quote'; $b.text = ($q.ToArray() -join "`n").Trim(); $blocks.Add($b)
            continue
        }
        if ([regex]::IsMatch($s, '^(-{3,}|\*{3,})$')) { $i++; continue }
        $para = New-Object System.Collections.Generic.List[string]
        while ($i -lt $n) {
            $s2 = $lines[$i].Trim()
            if ($s2 -ceq '' -or $s2.StartsWith('```', $script:Ord) -or $s2.StartsWith('<!--', $script:Ord) -or $s2.StartsWith('|', $script:Ord) -or
                $s2.StartsWith('>', $script:Ord) -or $s2.StartsWith('#', $script:Ord) -or [regex]::IsMatch($lines[$i], $script:ListRe)) { break }
            $para.Add($s2)
            $i++
        }
        $b = New-Block 'p'; $b.text = ($para.ToArray() -join "`n"); $blocks.Add($b)
    }
    return $doc
}
function Max-Units([string]$text) {
    $m = 0
    foreach ($x in (NormNl $text).Split([char]10)) { $u = TextUnits $x; if ($u -gt $m) { $m = $u } }
    return $m
}
function Section-Widths($blocks) {
    $tables = New-Object System.Collections.Generic.List[object]
    foreach ($b in $blocks) { if ($b.t -ceq 'table') { $tables.Add($b) } }
    $minTotal = [int]$script:Xl.min_total_width
    if ($tables.Count -eq 0) { return ,([int[]]@($minTotal)) }
    $n = 0
    foreach ($t in $tables) { if ($t.header.Length -gt $n) { $n = $t.header.Length } }
    $w = New-Object int[] $n
    foreach ($t in $tables) {
        $k = $t.header.Length
        for ($j = 0; $j -lt $k; $j++) {
            if ($j -eq $k - 1 -and $k -lt $n) { continue }
            if ($null -ne $t.cols -and $j -lt $t.cols.Length) {
                $cw = [int]$t.cols[$j]
            } else {
                $units = 0
                $all = New-Object System.Collections.Generic.List[object]
                $all.Add($t.header)
                foreach ($row in $t.rows) { $all.Add($row) }
                foreach ($row in $all) {
                    $cell = ''
                    if ($j -lt $row.Length) { $cell = $row[$j] }
                    $u = Max-Units (PlainOf (Md-Inline $cell))
                    if ($u -gt $units) { $units = $u }
                }
                $cw = [int][Math]::Min(60, [Math]::Max(8, [Math]::Floor(($units + 19) / 20) + 2))
            }
            if ($cw -gt $w[$j]) { $w[$j] = $cw }
        }
    }
    for ($j = 0; $j -lt $n; $j++) { if ($w[$j] -eq 0) { $w[$j] = 10 } }
    $total = 0
    foreach ($x in $w) { $total += $x }
    if ($total -lt $minTotal) { $w[$n - 1] += $minTotal - $total }
    return ,$w
}
function Add-Prefix([string]$prefix, $v) {
    if ($v -is [string]) { return $prefix + $v }
    $out = New-Object System.Collections.Generic.List[object]
    $arr = @($v)
    if ($arr.Count -gt 0 -and -not [bool]$arr[0][1]) {
        $out.Add(@(($prefix + [string]$arr[0][0]), $false))
        for ($k = 1; $k -lt $arr.Count; $k++) { $out.Add(@([string]$arr[$k][0], [bool]$arr[$k][1])) }
    } else {
        $out.Add(@($prefix, $false))
        foreach ($x in $arr) { $out.Add(@([string]$x[0], [bool]$x[1])) }
    }
    return ,($out.ToArray())
}
function Sum-Range($widths, [int]$from, [int]$to) {
    $s = 0
    for ($k = $from; $k -le $to; $k++) { $s += $widths[$k] }
    return $s
}
function Render-Table($sh, $b, [int]$r, [int]$first, [int]$last, $widths) {
    $header = $b.header
    $ncol = $header.Length
    $lines = 1
    for ($j = 0; $j -lt $ncol; $j++) {
        $c1 = $first + $j
        $c2 = $c1
        if ($j -eq $ncol - 1) { $c2 = $last }
        $w = Sum-Range $widths ($c1 - $first) ($c2 - $first)
        $v = Md-Inline $header[$j]
        Put-Block $sh $r $c1 $c2 $v 'th'
        $l = WrapLines (PlainOf $v) $w
        if ($l -gt $lines) { $lines = $l }
    }
    $sh.H10[$r] = RowH10 $lines
    $r++
    $k = 0
    foreach ($row in $b.rows) {
        $style = 'td'
        if ($k % 2 -eq 1) { $style = 'td_alt' }
        $lines = 1
        for ($j = 0; $j -lt $ncol; $j++) {
            $c1 = $first + $j
            $c2 = $c1
            if ($j -eq $ncol - 1) { $c2 = $last }
            $w = Sum-Range $widths ($c1 - $first) ($c2 - $first)
            $cell = ''
            if ($j -lt $row.Length) { $cell = $row[$j] }
            $v = Md-Inline $cell
            Put-Block $sh $r $c1 $c2 $v $style
            $l = WrapLines (PlainOf $v) $w
            if ($l -gt $lines) { $lines = $l }
        }
        $sh.H10[$r] = RowH10 $lines
        $r++
        $k++
    }
    return $r
}
function Render-Blocks($sh, $blocks, [int]$r, [int]$first, [int]$last, $widths) {
    $total = 0
    foreach ($x in $widths) { $total += $x }
    $k = 0
    foreach ($b in $blocks) {
        $t = $b.t
        if ($t -ceq 'h3') {
            if ($k -gt 0) { $sh.H10[$r] = 80; $r++ }
            Put-Block $sh $r $first $last (PlainOf (Md-Inline $b.text)) 'h3'
            $sh.H10[$r] = 240; $r++
            $sh.H10[$r] = 40; $r++
            $k++
            continue
        }
        if ($t -ceq 'h4') {
            Put-Block $sh $r $first $last (PlainOf (Md-Inline $b.text)) 'h4'
            $sh.H10[$r] = 220; $r++
            $k++
            continue
        }
        if ($t -ceq 'p') {
            $v = Md-Inline $b.text
            Put-Block $sh $r $first $last $v 'p'
            $sh.H10[$r] = RowH10 (WrapLines (PlainOf $v) $total)
            $r++
        } elseif ($t -ceq 'quote') {
            $v = Md-Inline $b.text
            Put-Block $sh $r $first $last $v 'quote'
            $sh.H10[$r] = RowH10 (WrapLines (PlainOf $v) ($total - 2))
            $r++
        } elseif ($t -ceq 'list') {
            foreach ($it in $b.items) {
                if ($it.mark -ceq '-' -or $it.mark -ceq '*') { $mark = '・' } else { $mark = $it.mark + ' ' }
                $v = Add-Prefix (('　' * $it.level) + $mark) (Md-Inline $it.text)
                Put-Block $sh $r $first $last $v 'li'
                $sh.H10[$r] = RowH10 (WrapLines (PlainOf $v) ($total - 2))
                $r++
            }
        } elseif ($t -ceq 'code') {
            $code = @($b.lines)
            if ($code.Count -eq 0) { $code = @('') }
            for ($j = 0; $j -lt $code.Count; $j++) {
                if ($code.Count -eq 1) { $style = 'code_single' }
                elseif ($j -eq 0) { $style = 'code_first' }
                elseif ($j -eq $code.Count - 1) { $style = 'code_last' }
                else { $style = 'code_mid' }
                $text = [string]$code[$j]
                if ($text -ceq '') { $text = ' ' }
                Put-Block $sh $r $first $last $text $style
                $sh.H10[$r] = RowH10 (WrapLines $text ($total - 2)) 18
                $r++
            }
        } elseif ($t -ceq 'table') {
            $r = Render-Table $sh $b $r $first $last $widths
        }
        $sh.H10[$r] = 60
        $r++
        $k++
    }
    return $r
}
function Render-Section([string]$name, [string]$title, $blocks) {
    $sh = New-Sheet $name
    $widths = Section-Widths $blocks
    $sh.Widths[1] = 2
    for ($j = 0; $j -lt $widths.Length; $j++) { $sh.Widths[$j + 2] = $widths[$j] }
    $first = 2
    $last = 1 + $widths.Length
    Put-Block $sh 1 $first $last $title 'title'
    $sh.H10[1] = 300
    $sh.H10[2] = 60
    [void](Render-Blocks $sh $blocks 3 $first $last $widths)
    return $sh
}
function Sheet-NameOf([string]$title, $used) {
    $name = [regex]::Replace($title, '[\\/?*\[\]:]', '').Trim()
    if ($name.Length -gt 31) { $name = $name.Substring(0, 31) }
    if (-not $name) { $name = 'Sheet' }
    $base = $name
    $k = 2
    while ($used.Contains($name.ToLowerInvariant())) {
        $suffix = '(' + $k + ')'
        $len = [Math]::Min($base.Length, 31 - $suffix.Length)
        $name = $base.Substring(0, $len) + $suffix
        $k++
    }
    [void]$used.Add($name.ToLowerInvariant())
    return $name
}
function Render-Cover($doc, $names) {
    $sh = New-Sheet $script:Xl.cover_sheet
    $sh.Selected = $true
    $widths = [int[]]@($script:Xl.cover_widths)
    $sh.Widths[1] = 2
    for ($j = 0; $j -lt $widths.Length; $j++) { $sh.Widths[$j + 2] = $widths[$j] }
    $first = 2
    $last = 1 + $widths.Length
    $sh.H10[1] = 120
    Put-Block $sh 2 $first $last $doc.title 'cover_title'
    $sh.H10[2] = 420
    $sh.H10[3] = 120
    $r = Render-Blocks $sh $doc.cover 4 $first $last $widths
    $sh.H10[$r] = 80; $r++
    Put-Block $sh $r $first $last $script:Xl.toc_title 'h3'
    $sh.H10[$r] = 240; $r++
    $sh.H10[$r] = 40; $r++
    $j = 0
    foreach ($h in $script:Xl.toc_headers) { Put-Cell $sh $r ($first + $j) $h 'th'; $j++ }
    $sh.H10[$r] = RowH10 1
    $r++
    for ($k = 0; $k -lt $doc.sections.Count; $k++) {
        $sec = $doc.sections[$k]
        Put-Cell $sh $r $first ([string]($k + 1)) 'toc_no'
        Put-Cell $sh $r ($first + 1) $names[$k] 'toc_link'
        $sh.Links.Add(@((CRef $r ($first + 1)), $names[$k]))
        Put-Cell $sh $r ($first + 2) $sec.summary 'toc_td'
        $sh.H10[$r] = RowH10 (WrapLines $sec.summary $widths[2])
        $r++
    }
    return $sh
}
function Md-Cell([string]$s) { return $s.Replace('|', '\|').Replace("`n", '<br>') }
function Md-Code([string]$s) {
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($x in (NormNl $s).Split([char]10)) { if ($x -cne '') { $parts.Add('`' + $x.Replace('|', '\|') + '`') } }
    return ($parts.ToArray() -join '<br>')
}
function Settings-TableMd {
    $out = New-Object System.Collections.Generic.List[string]
    $out.Add('<!-- cols: 12,28,22,16,62 -->')
    $out.Add('| 区分 | 設定キー | 設定項目 | 既定値 | 説明 |')
    $out.Add('|---|---|---|---|---|')
    foreach ($s in $script:Spec.settings) {
        $d = [string]$s.default
        if ($d -ceq '') { $d = $script:Tpl.empty_default }
        $desc = [string]$s.desc
        if (Has-Prop $s 'template') { $desc += '（テンプレートの記入例: ' + ([string]$s.template).Replace("`n", ' / ') + '）' }
        $out.Add('| ' + (Md-Cell $s.category) + ' | ' + (Md-Code $s.key) + ' | ' + (Md-Cell $s.label) + ' | ' + (Md-Cell $d) + ' | ' + (Md-Cell $desc) + ' |')
    }
    return ($out.ToArray() -join "`n")
}
function ApiColumns-TableMd {
    $out = New-Object System.Collections.Generic.List[string]
    $out.Add('<!-- cols: 18,6,62,36 -->')
    $out.Add('| 列名 | 必須 | 説明 | 記入例 |')
    $out.Add('|---|---|---|---|')
    foreach ($c in $script:Spec.api_columns) {
        $req = ''
        if ([bool]$c.required) { $req = '○' }
        $out.Add('| ' + (Md-Cell $c.header) + ' | ' + $req + ' | ' + (Md-Cell $c.desc) + ' | ' + (Md-Code $c.example) + ' |')
    }
    return ($out.ToArray() -join "`n")
}
function Expand-Md([string]$md) {
    $md = $md.Replace('1.0.0', $script:ToolVersion).Replace('@@DATE@@', (Get-Today))
    $md = $md.Replace('<!-- @SETTINGS_TABLE@ -->', (Settings-TableMd))
    return $md.Replace('<!-- @API_COLUMNS_TABLE@ -->', (ApiColumns-TableMd))
}

# ----------------------------------------------------------------------------- templates / data / manual
function Get-ApiValue($api, [string]$key) { return [string](Get-Prop $api $key '') }
function Template-ApiRows {
    $rows = New-Object System.Collections.Generic.List[object]
    $h = New-Object System.Collections.Generic.List[string]
    $hint = New-Object System.Collections.Generic.List[string]
    foreach ($c in $script:Spec.api_columns) { $h.Add($c.header); $hint.Add($c.hint) }
    $rows.Add($h.ToArray())
    $rows.Add($hint.ToArray())
    foreach ($api in $script:Spec.sample_apis) {
        $r = New-Object System.Collections.Generic.List[string]
        foreach ($c in $script:Spec.api_columns) { $r.Add((Get-ApiValue $api $c.key)) }
        $rows.Add($r.ToArray())
    }
    return ,$rows
}
function Setting-TemplateValue($s) {
    if (Has-Prop $s 'template') { return [string]$s.template }
    return [string]$s.default
}
function Template-SettingsRows($items = $null) {
    if ($null -eq $items) { $items = $script:Spec.settings }
    $rows = New-Object System.Collections.Generic.List[object]
    $rows.Add([string[]]@($script:Tpl.settings_headers))
    $cat = $null
    foreach ($s in $items) {
        if ($s.category -cne $cat) { $cat = $s.category; $rows.Add([string[]]@(('#■ ' + $cat), '', '', '', '')) }
        $d = [string]$s.default
        if ($d -ceq '') { $d = $script:Tpl.empty_default }
        $rows.Add([string[]]@($s.key, $s.label, (Setting-TemplateValue $s), $d, $s.desc))
    }
    return ,$rows
}
function Build-TemplateXlsx([string]$path) {
    $cols = @($script:Spec.api_columns)
    $n = $cols.Count
    $sh = New-Sheet $script:Tpl.api_sheet
    $sh.Selected = $true
    $total = 0
    for ($j = 0; $j -lt $n; $j++) { $sh.Widths[$j + 1] = [int]$cols[$j].width; $total += [int]$cols[$j].width }
    Put-Block $sh 1 1 $n $script:Tpl.api_title 'title'
    $sh.H10[1] = 300
    Put-Block $sh 2 1 $n $script:Tpl.api_note 'note'
    $sh.H10[2] = RowH10 (WrapLines $script:Tpl.api_note $total)
    $lines = 1
    for ($j = 0; $j -lt $n; $j++) {
        $style = 'api_hdr'
        if ([bool]$cols[$j].required) { $style = 'api_hdr_req' }
        Put-Cell $sh 3 ($j + 1) $cols[$j].header $style
        $l = WrapLines $cols[$j].header ([int]$cols[$j].width)
        if ($l -gt $lines) { $lines = $l }
    }
    $sh.H10[3] = RowH10 $lines
    $lines = 1
    for ($j = 0; $j -lt $n; $j++) {
        Put-Cell $sh 4 ($j + 1) $cols[$j].hint 'hint'
        $l = WrapLines $cols[$j].hint ([int]$cols[$j].width)
        if ($l -gt $lines) { $lines = $l }
    }
    $sh.H10[4] = RowH10 $lines
    $r = 5
    foreach ($api in $script:Spec.sample_apis) {
        $lines = 1
        for ($j = 0; $j -lt $n; $j++) {
            $v = Get-ApiValue $api $cols[$j].key
            Put-Cell $sh $r ($j + 1) $v 'cell'
            $l = WrapLines $v ([int]$cols[$j].width)
            if ($l -gt $lines) { $lines = $l }
        }
        $sh.H10[$r] = RowH10 $lines
        $r++
    }
    for ($k = 0; $k -lt [int]$script:Tpl.blank_rows; $k++) {
        for ($j = 0; $j -lt $n; $j++) { Put-Cell $sh $r ($j + 1) $null 'cell' }
        $sh.H10[$r] = RowH10 1
        $r++
    }
    $sh.FreezeRows = 4
    $sh.FreezeCols = 3
    $sh.Paper = 8
    $sh.FitWidth = 2
    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($c in $cols) { $keys.Add($c.key) }
    $last = [int]$script:Tpl.validation_rows
    foreach ($pair in @(@('enabled', 'Y,N'), @('method', 'GET,POST,PUT,PATCH,DELETE'), @('body_type', 'auto,json,form,raw'))) {
        $cl = ColLetter ($keys.IndexOf($pair[0]) + 1)
        $sh.Validations.Add(@(($cl + '5:' + $cl + $last), $pair[1]))
    }
    $st = New-Sheet $script:Tpl.settings_sheet
    $widths = [int[]]@($script:Tpl.settings_widths)
    $wsum = 0
    for ($j = 0; $j -lt $widths.Length; $j++) { $st.Widths[$j + 1] = $widths[$j]; $wsum += $widths[$j] }
    Put-Block $st 1 1 5 $script:Tpl.settings_title 'title'
    $st.H10[1] = 300
    Put-Block $st 2 1 5 $script:Tpl.settings_note 'note'
    $st.H10[2] = RowH10 (WrapLines $script:Tpl.settings_note $wsum)
    $j = 0
    foreach ($h in $script:Tpl.settings_headers) { Put-Cell $st 3 ($j + 1) $h 'th'; $j++ }
    $st.H10[3] = RowH10 1
    $r = 4
    $cat = $null
    foreach ($s in $script:Spec.settings) {
        if ($s.category -cne $cat) {
            $cat = $s.category
            Put-Block $st $r 1 5 ('■ ' + $cat) 'cat'
            $st.H10[$r] = 220
            $r++
        }
        $d = [string]$s.default
        if ($d -ceq '') { $d = $script:Tpl.empty_default }
        $cells = @(@($s.key, 'key'), @($s.label, 'key'), @((Setting-TemplateValue $s), 'input'), @($d, 'key'), @($s.desc, 'td'))
        $lines = 1
        for ($j = 0; $j -lt 5; $j++) {
            Put-Cell $st $r ($j + 1) ([string]$cells[$j][0]) ([string]$cells[$j][1])
            $l = WrapLines ([string]$cells[$j][0]) $widths[$j]
            if ($l -gt $lines) { $lines = $l }
        }
        $st.H10[$r] = RowH10 $lines
        if (Has-Prop $s 'choices') { $st.Validations.Add(@(('C' + $r), (@($s.choices) -join ','))) }
        $r++
    }
    $st.FreezeRows = 3
    $gdoc = Md-Parse (Expand-Md $script:Spec.template_guide_md)
    $sec = $gdoc.sections[0]
    $guide = Render-Section $script:Tpl.guide_sheet $sec.title $sec.blocks
    Save-Book @($sh, $st, $guide) $path $script:Tpl.api_title
}
function Build-TestdataCsv([string]$path) {
    $d = $script:Spec.sample_csv
    $rows = New-Object System.Collections.Generic.List[object]
    $rows.Add([string[]]@($d.header))
    foreach ($r in $d.rows) { $rows.Add([string[]]@($r)) }
    Write-TextFile $path (To-Csv $rows) $true
}
function Build-TestdataXlsx([string]$path) {
    $d = $script:Spec.sample_excel
    $sh = New-Sheet $d.sheet
    $sh.Selected = $true
    $j = 0
    foreach ($h in $d.header) { $sh.Widths[$j + 1] = 18; Put-Cell $sh 1 ($j + 1) $h 'data_hdr'; $j++ }
    $sh.H10[1] = 220
    $i = 0
    foreach ($row in $d.rows) {
        $j = 0
        foreach ($v in $row) { Put-Cell $sh ($i + 2) ($j + 1) ([string]$v) 'data_cell'; $j++ }
        $sh.H10[$i + 2] = 200
        $i++
    }
    $sh.FreezeRows = 1
    Save-Book @($sh) $path 'testdata'
}
function Build-Manual([string]$outDir) {
    $md = Expand-Md $script:Spec.manual_md
    $base = Join-Path $outDir $script:Xl.manual_file
    Write-TextFile ($base + '.md') $md
    $doc = Md-Parse $md
    $used = New-Object 'System.Collections.Generic.HashSet[string]'
    [void]$used.Add(([string]$script:Xl.cover_sheet).ToLowerInvariant())
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($sec in $doc.sections) { $names.Add((Sheet-NameOf $sec.title $used)) }
    $sheets = New-Object System.Collections.Generic.List[object]
    $sheets.Add((Render-Cover $doc $names))
    for ($k = 0; $k -lt $doc.sections.Count; $k++) {
        $sec = $doc.sections[$k]
        $sheets.Add((Render-Section $names[$k] $sec.title $sec.blocks))
    }
    Save-Book $sheets ($base + '.xlsx') $doc.title
    return @(($base + '.xlsx'), ($base + '.md'))
}

# ----------------------------------------------------------------------------- OpenAPI loading
# =============================================================================
#  OpenAPI 3.0.x 定義（JSON）の読み込み
#  JMeter_Tools の windows/OpenApi2Jmx.ps1 から、定義の読み込み・テストデータ生成の処理を
#  そのまま移植したものです（JMX 出力の部分は除く）。移植時の変更点は次のとおりです。
#    - Get-Prop を Get-OaProp に、$script:Warnings を $script:OaWarnings に改名した
#      （このツールの Get-Prop・$script:Warnings と名前が重なるため）
#    - New-OperationRequest の「6. 組み立て」を、API一覧（Excel / CSV）の 1 行と同じ項目を返すように変更した
#    - ヘッダ API-Key の警告文（このツールには共通の API-Key が無いため）
#  末尾の Get-OaRowPairs / Get-OaServerUrl / Load-OpenApi がこのツール用に追加した処理です。
#  RHEL 版 jmeter_api_tool.sh の同じ処理と同一の乱数列・同一の結果になります。
# =============================================================================
class JNum {
    [string]$Text
    JNum([string]$t) { $this.Text = $t }
}

$script:SKIP = New-Object System.Object

function Stop-Gen([string]$Message) {
    $ex = New-Object System.InvalidOperationException($Message)
    $ex.Data['OA2J'] = $true
    throw $ex
}

# ----------------------------------------------------------------- 乱数（両実装共通）
$script:RngState = 1L

function Initialize-Rng([long]$SeedValue) {
    $script:RngState = ($SeedValue % 2147483646L) + 1L
}

function Get-Next31 {
    $script:RngState = ($script:RngState * 48271L) % 2147483647L
    return $script:RngState
}

function Get-Between([long]$Lo, [long]$Hi) {
    if ($Hi -le $Lo) { return $Lo }
    $span = $Hi - $Lo + 1L
    if ($span -le 2147483646L) {
        $n = Get-Next31
        return [long]($Lo + (($n - 1L) % $span))
    }
    $a = (Get-Next31) - 1L
    $b = (Get-Next31) - 1L
    $r = $a * 2147483646L + $b
    return [long]($Lo + ($r % $span))
}

# ----------------------------------------------------------------- 値の判定・取得
$script:UNITS_LIMIT = 1000000000000000000L
$script:MAX_DECIMALS = 12
$script:MAX_DEPTH = 12
$script:ALNUM = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
$script:LOWER_ALNUM = 'abcdefghijklmnopqrstuvwxyz0123456789'
$script:HEXD = '0123456789abcdef'
$script:EPOCH_START = 1735689600L
$script:EPOCH_END = 1798761599L

function Test-Obj($x) { return ($x -is [System.Collections.IDictionary]) }
function Test-Arr($x) { return (($x -is [System.Collections.IList]) -and -not ($x -is [string])) }
function Test-Str($x) { return ($x -is [string]) }
function Test-Bool($x) { return ($x -is [bool]) }
function Test-Int($x) { return (($x -is [long]) -or ($x -is [int])) }
function Test-Dec($x) { return ($x -is [decimal]) }
function Test-Num($x) { return ((Test-Int $x) -or (Test-Dec $x)) }
function Test-Skip($x) { return [object]::ReferenceEquals($x, $script:SKIP) }

function Test-Has($o, [string]$k) {
    if ($o -is [System.Collections.Generic.Dictionary[string, object]]) { return $o.ContainsKey($k) }
    if ($o -is [System.Collections.Specialized.OrderedDictionary]) { return $o.Contains($k) }
    if ($o -is [System.Collections.IDictionary]) { return $o.Contains($k) }
    return $false
}

function Get-OaProp($o, [string]$k) {
    if (Test-Has $o $k) { return , ($o[$k]) }
    return $null
}

function New-OMap { return , (New-Object System.Collections.Specialized.OrderedDictionary) }
function New-OList { return , (New-Object 'System.Collections.Generic.List[object]') }

function Copy-OMap($src) {
    $r = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($k in @($src.Keys)) { $r[[string]$k] = $src[$k] }
    return , $r
}

function ConvertTo-Dec($x) {
    if (Test-Int $x) { return [decimal]$x }
    if (Test-Dec $x) { return $x }
    return $null
}

function Format-Dec([decimal]$d) { return $d.ToString($script:Inv) }

function Limit-DecTo([decimal]$d, [decimal]$Lim) {
    if ($d -gt $Lim) { return $Lim }
    if ($d -lt (-$Lim)) { return (-$Lim) }
    return $d
}

function Get-IntOrNull($x) {
    if (Test-Int $x) { if ($x -ge 0) { return [long]$x } else { return $null } }
    if ((Test-Dec $x) -and ([decimal]::Truncate($x) -eq $x) -and ($x -ge 0)) { return [long]$x }
    return $null
}

function Get-Pow10([int]$d) {
    $p = 1L
    for ($i = 0; $i -lt $d; $i++) { $p = $p * 10L }
    return $p
}

function Get-FloorDiv([long]$a, [long]$b) {
    $rem = 0L
    $q = [System.Math]::DivRem($a, $b, [ref]$rem)
    if (($rem -ne 0) -and (($rem -lt 0) -ne ($b -lt 0))) { $q = $q - 1L }
    return [long]$q
}

function Get-CeilDiv([long]$a, [long]$b) { return [long](-(Get-FloorDiv (-$a) $b)) }

function ConvertTo-GenValue($x) {
    if ($null -eq $x) { return $null }
    if ($x -is [bool]) { return $x }
    if (Test-Int $x) { return [JNum]::new(([long]$x).ToString($script:Inv)) }
    if (Test-Dec $x) { return [JNum]::new((Format-Dec $x)) }
    if ($x -is [string]) { return $x }
    if (Test-Arr $x) {
        $l = New-Object 'System.Collections.Generic.List[object]'
        foreach ($i in $x) { $l.Add((ConvertTo-GenValue $i)) }
        return , $l
    }
    if (Test-Obj $x) {
        $o = New-Object System.Collections.Specialized.OrderedDictionary
        foreach ($k in @($x.Keys)) { $o[[string]$k] = (ConvertTo-GenValue $x[$k]) }
        return , $o
    }
    return [string]$x
}

# ----------------------------------------------------------------- 文字列の整形
function Get-CleanText([string]$s) {
    if ($null -eq $s) { return '' }
    $sb = New-Object System.Text.StringBuilder
    $n = $s.Length
    for ($i = 0; $i -lt $n; $i++) {
        $c = [int]$s[$i]
        if (($c -ge 0xD800) -and ($c -le 0xDBFF)) {
            if (($i + 1 -lt $n) -and ([int]$s[$i + 1] -ge 0xDC00) -and ([int]$s[$i + 1] -le 0xDFFF)) {
                [void]$sb.Append($s[$i]).Append($s[$i + 1])
                $i++
            }
            continue
        }
        if (($c -ge 0xDC00) -and ($c -le 0xDFFF)) { continue }
        if (($c -eq 9) -or ($c -eq 10) -or ($c -eq 13) -or (($c -ge 0x20) -and ($c -le 0xD7FF)) -or (($c -ge 0xE000) -and ($c -le 0xFFFD))) {
            [void]$sb.Append($s[$i])
        }
    }
    return $sb.ToString()
}

function Get-JsonStr([string]$s) {
    $s = Get-CleanText $s
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('"')
    foreach ($ch in $s.ToCharArray()) {
        switch -CaseSensitive ($ch) {
            '"' { [void]$sb.Append('\"'); break }
            '\' { [void]$sb.Append('\\'); break }
            "`n" { [void]$sb.Append('\n'); break }
            "`r" { [void]$sb.Append('\r'); break }
            "`t" { [void]$sb.Append('\t'); break }
            default { [void]$sb.Append($ch) }
        }
    }
    [void]$sb.Append('"')
    return $sb.ToString()
}

function ConvertTo-JsonText($v, [int]$Indent, [bool]$Pretty) {
    if ($null -eq $v) { return 'null' }
    if ($v -is [bool]) { if ($v) { return 'true' } else { return 'false' } }
    if ($v -is [JNum]) { return $v.Text }
    if ($v -is [string]) { return (Get-JsonStr $v) }
    if (Test-Obj $v) {
        if ($v.Count -eq 0) { return '{}' }
        $parts = New-Object 'System.Collections.Generic.List[string]'
        if (-not $Pretty) {
            foreach ($k in @($v.Keys)) { $parts.Add((Get-JsonStr ([string]$k)) + ':' + (ConvertTo-JsonText $v[$k] 0 $false)) }
            return '{' + [string]::Join(',', $parts) + '}'
        }
        $pad = '  ' * ($Indent + 1)
        foreach ($k in @($v.Keys)) { $parts.Add($pad + (Get-JsonStr ([string]$k)) + ': ' + (ConvertTo-JsonText $v[$k] ($Indent + 1) $true)) }
        return "{`n" + [string]::Join(",`n", $parts) + "`n" + ('  ' * $Indent) + '}'
    }
    if (Test-Arr $v) {
        if ($v.Count -eq 0) { return '[]' }
        $parts = New-Object 'System.Collections.Generic.List[string]'
        if (-not $Pretty) {
            foreach ($x in $v) { $parts.Add((ConvertTo-JsonText $x 0 $false)) }
            return '[' + [string]::Join(',', $parts) + ']'
        }
        $pad = '  ' * ($Indent + 1)
        foreach ($x in $v) { $parts.Add($pad + (ConvertTo-JsonText $x ($Indent + 1) $true)) }
        return "[`n" + [string]::Join(",`n", $parts) + "`n" + ('  ' * $Indent) + ']'
    }
    return (Get-JsonStr ([string]$v))
}

function Get-ScalarStr($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [bool]) { if ($v) { return 'true' } else { return 'false' } }
    if ($v -is [JNum]) { return $v.Text }
    if ($v -is [string]) { return $v }
    return (ConvertTo-JsonText $v 0 $false)
}

function Get-PctEncode([string]$s, [bool]$KeepSlash) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes((Get-CleanText $s))
    $sb = New-Object System.Text.StringBuilder
    foreach ($b in $bytes) {
        $isUnreserved = (($b -ge 0x41) -and ($b -le 0x5A)) -or (($b -ge 0x61) -and ($b -le 0x7A)) -or (($b -ge 0x30) -and ($b -le 0x39)) -or ($b -eq 0x2D) -or ($b -eq 0x2E) -or ($b -eq 0x5F) -or ($b -eq 0x7E)
        if ($isUnreserved -or ($KeepSlash -and ($b -eq 0x2F))) {
            [void]$sb.Append([char]$b)
        } else {
            [void]$sb.Append('%').Append($b.ToString('X2'))
        }
    }
    return $sb.ToString()
}

function Get-HeaderSafe([string]$s) {
    return (Get-CleanText $s).Replace("`r", ' ').Replace("`n", ' ').Replace("`t", ' ')
}

function Get-SafeName([string]$s) {
    $s = [regex]::Replace($s, '[^A-Za-z0-9._-]+', '_')
    $s = [regex]::Replace($s, '_+', '_').Trim([char[]]@('_', '.'))
    if ($s.Length -eq 0) { return 'openapi' }
    return $s
}

function Get-CpLen([string]$s) {
    $n = 0
    foreach ($ch in $s.ToCharArray()) {
        $c = [int]$ch
        if (-not (($c -ge 0xDC00) -and ($c -le 0xDFFF))) { $n++ }
    }
    return $n
}

function ConvertFrom-CodePoint([int]$c) { return [char]::ConvertFromUtf32($c) }

# ----------------------------------------------------------------- 生成コンテキスト
$script:Doc = $null
$script:SeedValue = 0L
$script:OptRequiredOnly = $false
$script:OptUseExamples = $false
$script:OptAnyMethod = 'GET'
$script:OaWarnings = New-Object 'System.Collections.Generic.List[string]'

function Add-Warn([string]$m) {
    if (-not $script:OaWarnings.Contains($m)) { $script:OaWarnings.Add($m) }
}

function Get-UnescapedToken([string]$tok) {
    if ($tok.Contains('%')) {
        try { $tok = [System.Uri]::UnescapeDataString($tok) } catch { }
    }
    return $tok.Replace('~1', '/').Replace('~0', '~')
}

function Resolve-Ref($ref) {
    if (-not (Test-Str $ref)) { return $null }
    if (-not $ref.StartsWith('#', [System.StringComparison]::Ordinal)) { return $null }
    if (($ref -ceq '#') -or ($ref -ceq '#/')) { return , $script:Doc }
    if (-not $ref.StartsWith('#/', [System.StringComparison]::Ordinal)) { return $null }
    $cur = $script:Doc
    foreach ($tok in $ref.Substring(2).Split([char]'/')) {
        $key = Get-UnescapedToken $tok
        if ((Test-Obj $cur) -and (Test-Has $cur $key)) {
            $cur = $cur[$key]
        } elseif ((Test-Arr $cur) -and ($key -cmatch '^(0|[1-9][0-9]*)$') -and ([long]$key -lt $cur.Count)) {
            $cur = $cur[[int]$key]
        } else {
            return $null
        }
    }
    return , $cur
}

function Get-RefText($ref) {
    if (Test-Str $ref) { return $ref }
    return '(不正な $ref)'
}

function Resolve-Deref($obj, [string]$What) {
    $seen = New-Object 'System.Collections.Generic.List[string]'
    while ((Test-Obj $obj) -and (Test-Has $obj '$ref')) {
        $rt = Get-RefText $obj['$ref']
        if ($seen.Contains($rt) -or ($seen.Count -gt 32)) {
            Add-Warn ('{0} の $ref が循環しています: {1}' -f $What, $rt)
            return $null
        }
        $seen.Add($rt)
        $target = Resolve-Ref $obj['$ref']
        if ($null -eq $target) {
            Add-Warn ('{0} の $ref を解決できません（外部ファイル参照は未対応）: {1}' -f $What, $rt)
            return $null
        }
        $obj = $target
    }
    return , $obj
}

# ----------------------------------------------------------------- スキーマ → テストデータ
$script:TYPES = @('object', 'array', 'integer', 'number', 'boolean', 'string', 'null')

function Get-SchemaType($schema) {
    $t = Get-OaProp $schema 'type'
    if (Test-Arr $t) {
        $names = New-Object 'System.Collections.Generic.List[string]'
        foreach ($x in $t) { if (Test-Str $x) { $names.Add($x) } }
        $nonNull = New-Object 'System.Collections.Generic.List[string]'
        foreach ($x in $names) { if ($x -cne 'null') { $nonNull.Add($x) } }
        if ($nonNull.Count -gt 0) { $t = $nonNull[0] }
        elseif ($names.Count -gt 0) { $t = 'null' }
        else { $t = $null }
    }
    if ((Test-Str $t) -and ($script:TYPES -ccontains $t)) { return $t }
    if ((Test-Has $schema 'properties') -or (Test-Has $schema 'additionalProperties') -or (Test-Has $schema 'minProperties') -or (Test-Has $schema 'maxProperties')) { return 'object' }
    if ((Test-Has $schema 'items') -or (Test-Has $schema 'minItems') -or (Test-Has $schema 'maxItems')) { return 'array' }
    $fmt = Get-OaProp $schema 'format'
    if ((Test-Str $fmt) -and (($fmt -ceq 'int32') -or ($fmt -ceq 'int64'))) { return 'integer' }
    if ((Test-Str $fmt) -and (($fmt -ceq 'float') -or ($fmt -ceq 'double'))) { return 'number' }
    if ((Test-Has $schema 'minimum') -or (Test-Has $schema 'maximum') -or (Test-Has $schema 'multipleOf') -or (Test-Has $schema 'exclusiveMinimum') -or (Test-Has $schema 'exclusiveMaximum')) { return 'number' }
    return 'string'
}

function Test-ListContains($list, $value) {
    foreach ($x in $list) {
        if (($x -is [string]) -and ($value -is [string])) { if ($x -ceq $value) { return $true } }
        elseif ([object]::Equals($x, $value)) { return $true }
    }
    return $false
}

function Merge-Schemas($a, $b) {
    $r = Copy-OMap $a
    foreach ($k in @($b.Keys)) {
        $key = [string]$k
        $v = $b[$k]
        if (($key -ceq 'properties') -and (Test-Obj $v)) {
            $cur = Get-OaProp $r 'properties'
            if (Test-Obj $cur) { $props = Copy-OMap $cur } else { $props = New-OMap }
            foreach ($pk in @($v.Keys)) { if (-not $props.Contains([string]$pk)) { $props[[string]$pk] = $v[$pk] } }
            $r['properties'] = $props
        } elseif (($key -ceq 'required') -and (Test-Arr $v)) {
            $req = New-Object 'System.Collections.Generic.List[object]'
            $cur = Get-OaProp $r 'required'
            if (Test-Arr $cur) { foreach ($x in $cur) { $req.Add($x) } }
            foreach ($x in $v) { if (-not (Test-ListContains $req $x)) { $req.Add($x) } }
            $r['required'] = $req
        } elseif (-not $r.Contains($key)) {
            $r[$key] = $v
        }
    }
    return , $r
}

# 戻り値: @{ S = 実体 or $null; Stack = List[string]; State = 'ok'|'cycle'|'unresolved' }
function Resolve-SchemaRef($s, $stack) {
    $cur = $s
    $st = $stack
    while ((Test-Obj $cur) -and (Test-Has $cur '$ref')) {
        $rt = Get-RefText $cur['$ref']
        if ($st.Contains($rt)) { return @{ S = $null; Stack = $st; State = 'cycle' } }
        $t = Resolve-Ref $cur['$ref']
        if ($null -eq $t) {
            Add-Warn ('スキーマの $ref を解決できません（外部ファイル参照は未対応）: ' + $rt)
            return @{ S = $null; Stack = $st; State = 'unresolved' }
        }
        $ns = New-Object 'System.Collections.Generic.List[string]'
        foreach ($x in $st) { $ns.Add($x) }
        $ns.Add($rt)
        $st = $ns
        $cur = $t
    }
    return @{ S = $cur; Stack = $st; State = 'ok' }
}

function Merge-AllOf($schema, $stack) {
    $result = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($k in @($schema.Keys)) { if ([string]$k -cne 'allOf') { $result[[string]$k] = $schema[$k] } }
    foreach ($sub in $schema['allOf']) {
        $res = Resolve-SchemaRef $sub $stack
        if (($res.State -cne 'ok') -or -not (Test-Obj $res.S)) { continue }
        $s = $res.S
        if (Test-Arr (Get-OaProp $s 'allOf')) { $s = Merge-AllOf $s $res.Stack }
        $result = Merge-Schemas $result $s
    }
    return , $result
}

function New-Gen($schema, [int]$Depth, $stack) {
    if (($null -eq $schema) -or (($schema -is [bool]) -and $schema)) { $schema = New-OMap }
    if (($schema -is [bool]) -or -not (Test-Obj $schema)) { return $script:SKIP }
    if (Test-Has $schema '$ref') {
        $res = Resolve-SchemaRef $schema $stack
        if ($res.State -ceq 'cycle') { return $script:SKIP }
        if ($res.State -ceq 'unresolved') { return 'UNRESOLVED_REF' }
        return , (New-Gen $res.S $Depth $res.Stack)
    }
    if ($Depth -gt $script:MAX_DEPTH) { return $script:SKIP }
    if (Test-Arr (Get-OaProp $schema 'allOf')) { $schema = Merge-AllOf $schema $stack }
    foreach ($key in @('oneOf', 'anyOf')) {
        $alts = Get-OaProp $schema $key
        if ((Test-Arr $alts) -and ($alts.Count -gt 0)) {
            $base = New-Object System.Collections.Specialized.OrderedDictionary
            foreach ($k in @($schema.Keys)) { if (([string]$k -cne 'oneOf') -and ([string]$k -cne 'anyOf')) { $base[[string]$k] = $schema[$k] } }
            $res = Resolve-SchemaRef $alts[0] $stack
            if ($res.State -ceq 'cycle') { return $script:SKIP }
            if ($res.State -ceq 'unresolved') { return 'UNRESOLVED_REF' }
            $alt = $res.S
            if (-not (Test-Obj $alt)) { $alt = New-OMap }
            return , (New-Gen (Merge-Schemas $base $alt) $Depth $res.Stack)
        }
    }
    if ($script:OptUseExamples) {
        if (Test-Has $schema 'example') { return , (ConvertTo-GenValue $schema['example']) }
        $exs = Get-OaProp $schema 'examples'
        if ((Test-Arr $exs) -and ($exs.Count -gt 0)) { return , (ConvertTo-GenValue $exs[0]) }
        if (Test-Has $schema 'default') { return , (ConvertTo-GenValue $schema['default']) }
    }
    if (Test-Has $schema 'const') { return , (ConvertTo-GenValue $schema['const']) }
    $enum = Get-OaProp $schema 'enum'
    if ((Test-Arr $enum) -and ($enum.Count -gt 0)) {
        $cands = New-Object 'System.Collections.Generic.List[object]'
        foreach ($e in $enum) { if ($null -ne $e) { $cands.Add($e) } }
        if ($cands.Count -eq 0) { return $null }
        return , (ConvertTo-GenValue $cands[[int](Get-Between 0 ($cands.Count - 1))])
    }
    $t = Get-SchemaType $schema
    switch -CaseSensitive ($t) {
        'object' { return , (New-GenObject $schema $Depth $stack) }
        'array' { return , (New-GenArray $schema $Depth $stack) }
        'integer' { return (New-GenInteger $schema) }
        'number' { return (New-GenNumber $schema) }
        'boolean' { return ((Get-Between 0 1) -eq 1) }
        'null' { return $null }
    }
    return (New-GenString $schema)
}

function Test-ReadOnly($pschema) {
    $s = $pschema
    $guard = 0
    while ((Test-Obj $s) -and (Test-Has $s '$ref') -and ($guard -lt 32)) {
        $s = Resolve-Ref $s['$ref']
        $guard++
    }
    if (-not (Test-Obj $s)) { return $false }
    $ro = Get-OaProp $s 'readOnly'
    return (($ro -is [bool]) -and $ro)
}

function New-GenObject($schema, [int]$Depth, $stack) {
    $result = New-Object System.Collections.Specialized.OrderedDictionary
    $props = Get-OaProp $schema 'properties'
    $req = Get-OaProp $schema 'required'
    $required = New-Object 'System.Collections.Generic.List[string]'
    if (Test-Arr $req) { foreach ($x in $req) { if (Test-Str $x) { $required.Add($x) } } }
    if (Test-Obj $props) {
        foreach ($nk in @($props.Keys)) {
            $name = [string]$nk
            if ($script:OptRequiredOnly -and -not $required.Contains($name)) { continue }
            $pschema = $props[$nk]
            if (Test-ReadOnly $pschema) { continue }
            $v = New-Gen $pschema ($Depth + 1) $stack
            if (Test-Skip $v) { continue }
            $result[$name] = $v
        }
    }
    $addl = Get-OaProp $schema 'additionalProperties'
    $propsCount = 0
    if (Test-Obj $props) { $propsCount = $props.Count }
    if (($result.Count -eq 0) -and ($propsCount -eq 0) -and (Test-Obj $addl) -and ($addl.Count -gt 0) -and -not $script:OptRequiredOnly) {
        $v = New-Gen $addl ($Depth + 1) $stack
        if (-not (Test-Skip $v)) { $result['key1'] = $v }
    }
    return , $result
}

function New-GenArray($schema, [int]$Depth, $stack) {
    $items = Get-OaProp $schema 'items'
    if (-not (Test-Obj $items)) { $items = New-OMap }
    $mn = Get-IntOrNull (Get-OaProp $schema 'minItems')
    $mx = Get-IntOrNull (Get-OaProp $schema 'maxItems')
    $n = 0L
    if ($null -ne $mn) { $n = $mn }
    if ($n -lt 1) { $n = 1L }
    if (($null -ne $mx) -and ($n -gt $mx)) { $n = $mx }
    if ($n -gt 50) { $n = 50L }
    $uq = Get-OaProp $schema 'uniqueItems'
    $unique = (($uq -is [bool]) -and $uq)
    $result = New-Object 'System.Collections.Generic.List[object]'
    $seen = New-Object 'System.Collections.Generic.List[string]'
    for ($i = 0; $i -lt $n; $i++) {
        $v = New-Gen $items ($Depth + 1) $stack
        if (Test-Skip $v) { break }
        if ($unique) {
            $key = ConvertTo-JsonText $v 0 $false
            $tries = 0
            while ($seen.Contains($key) -and ($tries -lt 10)) {
                $v = New-Gen $items ($Depth + 1) $stack
                if (Test-Skip $v) { break }
                $key = ConvertTo-JsonText $v 0 $false
                $tries++
            }
            if (Test-Skip $v) { break }
            $seen.Add($key)
        }
        $result.Add($v)
    }
    return , $result
}

# 戻り値: @(lo, hi)（scale 単位の整数）
function Get-NumBounds($schema, [long]$Scale) {
    $mn = ConvertTo-Dec (Get-OaProp $schema 'minimum')
    $mx = ConvertTo-Dec (Get-OaProp $schema 'maximum')
    $exmin = Get-OaProp $schema 'exclusiveMinimum'
    $exmax = Get-OaProp $schema 'exclusiveMaximum'
    $exLo = (($exmin -is [bool]) -and $exmin)
    $exHi = (($exmax -is [bool]) -and $exmax)
    if (Test-Num $exmin) {
        $e = ConvertTo-Dec $exmin
        if (($null -eq $mn) -or ($e -ge $mn)) { $mn = $e; $exLo = $true }
    }
    if (Test-Num $exmax) {
        $e = ConvertTo-Dec $exmax
        if (($null -eq $mx) -or ($e -le $mx)) { $mx = $e; $exHi = $true }
    }
    $vlim = [decimal](Get-FloorDiv $script:UNITS_LIMIT $Scale)
    $lo = $null
    $hi = $null
    if ($null -ne $mn) {
        $v = (Limit-DecTo $mn $vlim) * [decimal]$Scale
        $lo = [long][decimal]::Ceiling($v)
        if ($exLo -and ([decimal]$lo -eq $v)) { $lo = $lo + 1L }
    }
    if ($null -ne $mx) {
        $v = (Limit-DecTo $mx $vlim) * [decimal]$Scale
        $hi = [long][decimal]::Floor($v)
        if ($exHi -and ([decimal]$hi -eq $v)) { $hi = $hi - 1L }
    }
    $one = $Scale
    if (($null -eq $lo) -and ($null -eq $hi)) {
        $lo = $one
        $hi = 100L * $one
    } elseif ($null -eq $hi) {
        $hi = $lo + 99L * $one
    } elseif ($null -eq $lo) {
        if ($hi -ge $one) { $lo = $one } else { $lo = $hi - 99L * $one }
    }
    if ($lo -gt $hi) { $hi = $lo }
    return @([long]$lo, [long]$hi)
}

function New-GenInteger($schema) {
    $b = Get-NumBounds $schema 1L
    $lo = [long]$b[0]
    $hi = [long]$b[1]
    $fmt = Get-OaProp $schema 'format'
    if ((Test-Str $fmt) -and ($fmt -ceq 'int32')) {
        if ($lo -lt -2147483648L) { $lo = -2147483648L }
        if ($hi -gt 2147483647L) { $hi = 2147483647L }
        if ($lo -gt $hi) { $hi = $lo }
    }
    $m = ConvertTo-Dec (Get-OaProp $schema 'multipleOf')
    if (($null -ne $m) -and ($m -gt 0) -and ([decimal]::Truncate($m) -eq $m)) {
        $mi = [long]$m
        $kmin = Get-CeilDiv $lo $mi
        $kmax = Get-FloorDiv $hi $mi
        if ($kmin -le $kmax) { return [JNum]::new(((Get-Between $kmin $kmax) * $mi).ToString($script:Inv)) }
        return [JNum]::new($lo.ToString($script:Inv))
    }
    return [JNum]::new((Get-Between $lo $hi).ToString($script:Inv))
}

function Get-DecimalsOf([decimal]$d) {
    $s = Format-Dec $d
    $i = $s.IndexOf('.')
    if ($i -ge 0) { return $s.Substring($i + 1).TrimEnd([char]'0').Length }
    return 0
}

function Format-Units([long]$Units, [int]$D) {
    $neg = ($Units -lt 0)
    if ($neg) { $a = -$Units } else { $a = $Units }
    if ($D -eq 0) {
        $s = $a.ToString($script:Inv)
    } else {
        $p = Get-Pow10 $D
        $rem = 0L
        $q = [System.Math]::DivRem([long]$a, [long]$p, [ref]$rem)
        $s = ([long]$q).ToString($script:Inv) + '.' + ([long]$rem).ToString($script:Inv).PadLeft($D, [char]'0')
    }
    if ($neg) { return '-' + $s }
    return $s
}

function New-GenNumber($schema) {
    $m = ConvertTo-Dec (Get-OaProp $schema 'multipleOf')
    if (($null -ne $m) -and ($m -le 0)) { $m = $null }
    if ($null -ne $m) { $d = Get-DecimalsOf $m } else { $d = 2 }
    if ($d -gt $script:MAX_DECIMALS) { $d = $script:MAX_DECIMALS }
    $scale = Get-Pow10 $d
    $b = Get-NumBounds $schema $scale
    $lo = [long]$b[0]
    $hi = [long]$b[1]
    $units = $null
    if ($null -ne $m) {
        $mu = $m * [decimal]$scale
        if (($mu -gt 0) -and ([decimal]::Truncate($mu) -eq $mu)) {
            $mi = [long]$mu
            $kmin = Get-CeilDiv $lo $mi
            $kmax = Get-FloorDiv $hi $mi
            if ($kmin -le $kmax) { $units = (Get-Between $kmin $kmax) * $mi } else { $units = $lo }
        }
    }
    if ($null -eq $units) { $units = Get-Between $lo $hi }
    return [JNum]::new((Format-Units ([long]$units) $d))
}

function Get-RandStr([string]$Pool, [long]$N) {
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $N; $i++) {
        [void]$sb.Append($Pool[[int](Get-Between 0 ($Pool.Length - 1))])
    }
    return $sb.ToString()
}

function Get-ChooseLen($minL, $maxL) {
    $lo = 0L
    if ($null -ne $minL) { $lo = [long]$minL }
    if ($lo -lt 8) { $lo = 8L }
    if (($null -ne $maxL) -and ($lo -gt $maxL)) { $lo = [long]$maxL }
    if ($null -eq $maxL) {
        $hi = [System.Math]::Max($lo, 16L)
    } else {
        $hi = [System.Math]::Min([long]$maxL, [System.Math]::Max($lo, 16L))
    }
    if ($hi -lt $lo) { $hi = $lo }
    return (Get-Between $lo $hi)
}

function Get-UtcFromEpoch([long]$Secs) {
    return ([datetime]::new(1970, 1, 1, 0, 0, 0, [System.DateTimeKind]::Utc)).AddSeconds([double]$Secs)
}

function New-GenString($schema) {
    $fmt = Get-OaProp $schema 'format'
    if (-not (Test-Str $fmt)) { $fmt = '' }
    $minL = Get-IntOrNull (Get-OaProp $schema 'minLength')
    $maxL = Get-IntOrNull (Get-OaProp $schema 'maxLength')
    $pattern = Get-OaProp $schema 'pattern'
    if ((Test-Str $pattern) -and ($pattern.Length -gt 0)) {
        $s = New-FromPattern $pattern $minL $maxL
        if ($null -ne $s) { return $s }
    }
    switch -CaseSensitive ($fmt) {
        'date-time' {
            return (Get-UtcFromEpoch (Get-Between $script:EPOCH_START $script:EPOCH_END)).ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'", $script:Inv)
        }
        'date' {
            $days = Get-Between ([long][System.Math]::Floor($script:EPOCH_START / 86400)) ([long][System.Math]::Floor($script:EPOCH_END / 86400))
            return (Get-UtcFromEpoch ($days * 86400L)).ToString("yyyy'-'MM'-'dd", $script:Inv)
        }
        'time' {
            $secs = Get-Between 0 86399
            $hh = Get-FloorDiv $secs 3600
            $mm = Get-FloorDiv ($secs % 3600L) 60
            $ss = $secs % 60L
            return ('{0}:{1}:{2}' -f ([long]$hh).ToString('00', $script:Inv), ([long]$mm).ToString('00', $script:Inv), ([long]$ss).ToString('00', $script:Inv))
        }
        'email' { return (Get-RandStr $script:LOWER_ALNUM 8) + '@example.com' }
        'uuid' {
            $sb = New-Object System.Text.StringBuilder
            for ($i = 0; $i -lt 32; $i++) {
                if ($i -eq 12) { [void]$sb.Append('4') }
                elseif ($i -eq 16) { [void]$sb.Append($script:HEXD[[int](8 + (Get-Between 0 3))]) }
                else { [void]$sb.Append($script:HEXD[[int](Get-Between 0 15)]) }
            }
            $u = $sb.ToString()
            return $u.Substring(0, 8) + '-' + $u.Substring(8, 4) + '-' + $u.Substring(12, 4) + '-' + $u.Substring(16, 4) + '-' + $u.Substring(20, 12)
        }
        'uri' { return 'https://example.com/' + (Get-RandStr $script:LOWER_ALNUM 8) }
        'url' { return 'https://example.com/' + (Get-RandStr $script:LOWER_ALNUM 8) }
        'uri-reference' { return 'https://example.com/' + (Get-RandStr $script:LOWER_ALNUM 8) }
        'iri' { return 'https://example.com/' + (Get-RandStr $script:LOWER_ALNUM 8) }
        'iri-reference' { return 'https://example.com/' + (Get-RandStr $script:LOWER_ALNUM 8) }
        'hostname' { return (Get-RandStr $script:LOWER_ALNUM 8) + '.example.com' }
        'idn-hostname' { return (Get-RandStr $script:LOWER_ALNUM 8) + '.example.com' }
        'ipv4' { return '192.0.2.' + (Get-Between 1 254).ToString($script:Inv) }
        'ipv6' { return '2001:db8::' + (Get-Between 1 65535).ToString('x') }
        'byte' {
            $raw = New-Object byte[] 12
            for ($i = 0; $i -lt 12; $i++) { $raw[$i] = [byte](Get-Between 0 255) }
            return [System.Convert]::ToBase64String($raw)
        }
        'password' { return (Get-RandStr $script:ALNUM 12) }
        'binary' { return (Get-RandStr $script:ALNUM 16) }
    }
    return (Get-RandStr $script:ALNUM (Get-ChooseLen $minL $maxL))
}

# ----------------------------------------------------------------- pattern（正規表現）から文字列を生成
function New-CodeRange([int]$a, [int]$b) {
    $l = New-Object 'System.Collections.Generic.List[int]'
    for ($c = $a; $c -le $b; $c++) { $l.Add($c) }
    return , $l
}

function New-SortedCodes($list) {
    $set = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($c in $list) { [void]$set.Add([int]$c) }
    $l = New-Object 'System.Collections.Generic.List[int]' (, [int[]]@($set))
    $l.Sort()
    return , ($l.ToArray())
}

$script:DIGIT_SET = (New-CodeRange 0x30 0x39).ToArray()
$script:UPPER_SET = (New-CodeRange 0x41 0x5A).ToArray()
$script:LOWER_SET = (New-CodeRange 0x61 0x7A).ToArray()
$script:ALNUM_SET = [int[]]($script:DIGIT_SET + $script:UPPER_SET + $script:LOWER_SET)
$script:WORD_SET = New-SortedCodes ([int[]]($script:ALNUM_SET + @(0x5F)))
$script:NONDIGIT_SET = [int[]]($script:UPPER_SET + $script:LOWER_SET)
$script:NONWORD_SET = [int[]]@(0x2D, 0x2E)
$script:SPACE_SET = [int[]]@(0x20)
$script:NEG_BASE = New-SortedCodes ([int[]]($script:ALNUM_SET + @(0x2D, 0x2E, 0x5F)))
$script:PRINTABLE_SET = (New-CodeRange 0x21 0x7E).ToArray()

function Test-ValidCode([int]$c) {
    return (($c -eq 9) -or ($c -eq 10) -or ($c -eq 13) -or (($c -ge 0x20) -and ($c -le 0xD7FF)) -or (($c -ge 0xE000) -and ($c -le 0xFFFD)) -or (($c -ge 0x10000) -and ($c -le 0x10FFFF)))
}

# パターン文字列はコードポイント単位の配列で扱う（Python の str と同じ単位）
$script:PP = $null
$script:PI = 0

function Stop-Pattern([string]$m) {
    $ex = New-Object System.FormatException($m)
    $ex.Data['PATTERN'] = $true
    throw $ex
}

function Get-PPeek([int]$k = 0) {
    $j = $script:PI + $k
    if ($j -lt $script:PP.Count) { return $script:PP[$j] }
    return $null
}

function Get-PTake {
    $ch = $script:PP[$script:PI]
    $script:PI++
    return $ch
}

function New-PNode([string]$T) { return @{ T = $T } }

function Read-PAlt {
    $options = New-Object 'System.Collections.Generic.List[object]'
    $options.Add((Read-PSeq))
    while ((Get-PPeek) -ceq '|') {
        [void](Get-PTake)
        $options.Add((Read-PSeq))
    }
    if ($options.Count -eq 1) { return $options[0] }
    $n = New-PNode 'alt'
    $n.Items = $options
    return $n
}

function Read-PSeq {
    $items = New-Object 'System.Collections.Generic.List[object]'
    while ($true) {
        $ch = Get-PPeek
        if (($null -eq $ch) -or ($ch -ceq '|') -or ($ch -ceq ')')) { break }
        $items.Add((Read-PQuant (Read-PAtom)))
    }
    $n = New-PNode 'seq'
    $n.Items = $items
    return $n
}

function Get-EmptySeq {
    $n = New-PNode 'seq'
    $n.Items = New-Object 'System.Collections.Generic.List[object]'
    return $n
}

function New-PLit([int]$c) { $n = New-PNode 'lit'; $n.C = $c; return $n }
function New-PSet($pool) { $n = New-PNode 'set'; $n.Pool = [int[]]$pool; return $n }

function Read-PAtom {
    $ch = Get-PTake
    if ($ch -ceq '(') {
        if ((Get-PPeek) -ceq '?') {
            [void](Get-PTake)
            $c2 = Get-PPeek
            if ($c2 -ceq ':') { [void](Get-PTake) }
            elseif (($c2 -ceq 'P') -and ((Get-PPeek 1) -ceq '<')) { Skip-PGroupName 2 }
            elseif (($c2 -ceq '<') -and ($null -ne (Get-PPeek 1)) -and ((Get-PPeek 1) -cne '=') -and ((Get-PPeek 1) -cne '!')) { Skip-PGroupName 1 }
            else { Stop-Pattern 'unsupported group' }
        }
        $node = Read-PAlt
        if ((Get-PPeek) -cne ')') { Stop-Pattern 'unbalanced' }
        [void](Get-PTake)
        return $node
    }
    if ($ch -ceq '[') { return (Read-PClass) }
    if ($ch -ceq '.') { return (New-PSet $script:ALNUM_SET) }
    if (($ch -ceq '^') -or ($ch -ceq '$')) { return (Get-EmptySeq) }
    if ($ch -ceq '\') { return (Read-PEscape $false) }
    if (($ch -ceq '*') -or ($ch -ceq '+') -or ($ch -ceq '?') -or ($ch -ceq ')')) { Stop-Pattern 'quantifier without target' }
    return (New-PLit ([char]::ConvertToUtf32($ch, 0)))
}

function Skip-PGroupName([int]$k) {
    for ($i = 0; $i -lt $k; $i++) { [void](Get-PTake) }
    while ($true) {
        $c = Get-PPeek
        if ($null -eq $c) { Stop-Pattern 'bad group name' }
        [void](Get-PTake)
        if ($c -ceq '>') { return }
    }
}

function Read-PHex([int]$Count) {
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Count; $i++) {
        $c = Get-PPeek $i
        if ($null -eq $c) { Stop-Pattern 'bad hex escape' }
        [void]$sb.Append($c)
    }
    $s = $sb.ToString()
    if (-not ($s -cmatch '^[0-9A-Fa-f]+$')) { Stop-Pattern 'bad hex escape' }
    $script:PI += $Count
    return [System.Convert]::ToInt32($s, 16)
}

function Read-PEscape([bool]$InClass) {
    if ($script:PI -ge $script:PP.Count) { Stop-Pattern 'trailing backslash' }
    $ch = Get-PTake
    switch -CaseSensitive ($ch) {
        'd' { return (New-PSet $script:DIGIT_SET) }
        'D' { return (New-PSet $script:NONDIGIT_SET) }
        'w' { return (New-PSet $script:WORD_SET) }
        'W' { return (New-PSet $script:NONWORD_SET) }
        's' { return (New-PSet $script:SPACE_SET) }
        'S' { return (New-PSet $script:ALNUM_SET) }
        't' { return (New-PLit 9) }
        'n' { return (New-PLit 10) }
        'r' { return (New-PLit 13) }
        'u' {
            $c = Read-PHex 4
            if (-not (Test-ValidCode $c)) { Stop-Pattern 'bad code point' }
            return (New-PLit $c)
        }
        'x' {
            $c = Read-PHex 2
            if (-not (Test-ValidCode $c)) { Stop-Pattern 'bad code point' }
            return (New-PLit $c)
        }
    }
    if (@('b', 'B', 'A', 'z', 'Z', 'G') -ccontains $ch) {
        if ($InClass) { Stop-Pattern 'unsupported escape in class' }
        return (Get-EmptySeq)
    }
    if ('0123456789'.Contains($ch) -or (@('p', 'P', 'k', 'c', 'f', 'v', 'e', 'a', 'Q', 'E') -ccontains $ch)) { Stop-Pattern 'unsupported escape' }
    return (New-PLit ([char]::ConvertToUtf32($ch, 0)))
}

function Read-PClassItem {
    $ch = Get-PTake
    if ($ch -ceq '\') { return (Read-PEscape $true) }
    if (($ch -ceq '[') -and ((Get-PPeek) -ceq ':')) { Stop-Pattern 'posix class' }
    return (New-PLit ([char]::ConvertToUtf32($ch, 0)))
}

function Read-PClass {
    $negate = $false
    if ((Get-PPeek) -ceq '^') { [void](Get-PTake); $negate = $true }
    $codes = New-Object 'System.Collections.Generic.HashSet[int]'
    $first = $true
    while ($true) {
        if ($script:PI -ge $script:PP.Count) { Stop-Pattern 'unterminated class' }
        if (((Get-PPeek) -ceq ']') -and -not $first) { [void](Get-PTake); break }
        $first = $false
        $item = Read-PClassItem
        if (($item.T -ceq 'lit') -and ((Get-PPeek) -ceq '-') -and ($null -ne (Get-PPeek 1)) -and ((Get-PPeek 1) -cne ']')) {
            [void](Get-PTake)
            $item2 = Read-PClassItem
            if ($item2.T -cne 'lit') { Stop-Pattern 'bad range' }
            $a = [int]$item.C
            $b = [int]$item2.C
            if (($a -gt $b) -or (($b - $a) -gt 70000)) { Stop-Pattern 'bad range' }
            for ($c = $a; $c -le $b; $c++) { [void]$codes.Add($c) }
        } elseif ($item.T -ceq 'lit') {
            [void]$codes.Add([int]$item.C)
        } elseif ($item.T -ceq 'set') {
            foreach ($c in $item.Pool) { [void]$codes.Add([int]$c) }
        } else {
            Stop-Pattern 'bad class item'
        }
    }
    $pool = New-Object 'System.Collections.Generic.List[int]'
    if ($negate) {
        foreach ($c in $script:NEG_BASE) { if (-not $codes.Contains($c)) { $pool.Add($c) } }
        if ($pool.Count -eq 0) {
            foreach ($c in $script:PRINTABLE_SET) { if (-not $codes.Contains($c)) { $pool.Add($c) } }
        }
    } else {
        foreach ($c in $codes) { $pool.Add($c) }
        $pool.Sort()
    }
    $valid = New-Object 'System.Collections.Generic.List[int]'
    foreach ($c in $pool) { if (Test-ValidCode $c) { $valid.Add($c) } }
    if ($valid.Count -eq 0) { Stop-Pattern 'empty class' }
    return (New-PSet $valid.ToArray())
}

function Read-PQuant($node) {
    while ($true) {
        $ch = Get-PPeek
        if ($ch -ceq '*') { [void](Get-PTake); $mn = 0; $mx = -1 }
        elseif ($ch -ceq '+') { [void](Get-PTake); $mn = 1; $mx = -1 }
        elseif ($ch -ceq '?') { [void](Get-PTake); $mn = 0; $mx = 1 }
        elseif ($ch -ceq '{') {
            $sb = New-Object System.Text.StringBuilder
            for ($j = $script:PI; $j -lt $script:PP.Count; $j++) { [void]$sb.Append($script:PP[$j]) }
            $m = [regex]::Match($sb.ToString(), '^\{([0-9]+)(,([0-9]*))?\}')
            if (-not $m.Success) { return $node }
            $script:PI += $m.Value.Length
            $mn = [int]::Parse($m.Groups[1].Value, $script:Inv)
            if (-not $m.Groups[2].Success) { $mx = $mn }
            elseif ($m.Groups[3].Value -ceq '') { $mx = -1 }
            else { $mx = [int]::Parse($m.Groups[3].Value, $script:Inv) }
            if ((($mx -ne -1) -and ($mx -lt $mn)) -or ($mn -gt 1000)) { Stop-Pattern 'bad quantifier' }
        } else {
            return $node
        }
        if (((Get-PPeek) -ceq '?') -or ((Get-PPeek) -ceq '+')) { [void](Get-PTake) }
        $r = New-PNode 'rep'
        $r.Node = $node
        $r.Min = $mn
        $r.Max = $mx
        $node = $r
    }
}

function Write-PNode($node, [System.Text.StringBuilder]$Out) {
    switch -CaseSensitive ($node.T) {
        'lit' { [void]$Out.Append((ConvertFrom-CodePoint $node.C)) }
        'set' {
            $pool = $node.Pool
            [void]$Out.Append((ConvertFrom-CodePoint $pool[[int](Get-Between 0 ($pool.Length - 1))]))
        }
        'seq' { foreach ($x in $node.Items) { Write-PNode $x $Out } }
        'alt' {
            $opts = $node.Items
            Write-PNode $opts[[int](Get-Between 0 ($opts.Count - 1))] $Out
        }
        'rep' {
            $mn = [long]$node.Min
            $mx = [long]$node.Max
            if ($mx -eq -1) { $mx = $mn + 5 }
            $cap = [System.Math]::Max($mn, 16L)
            if ($mx -gt $cap) { $mx = $cap }
            $k = Get-Between $mn $mx
            for ($i = 0; $i -lt $k; $i++) { Write-PNode $node.Node $Out }
        }
    }
}

function ConvertTo-CodePointList([string]$s) {
    $l = New-Object 'System.Collections.Generic.List[string]'
    $i = 0
    while ($i -lt $s.Length) {
        if ([char]::IsHighSurrogate($s[$i]) -and ($i + 1 -lt $s.Length) -and [char]::IsLowSurrogate($s[$i + 1])) {
            $l.Add($s.Substring($i, 2))
            $i += 2
        } else {
            $l.Add([string]$s[$i])
            $i++
        }
    }
    return , $l
}

function New-FromPattern([string]$Pattern, $minL, $maxL) {
    $script:PP = ConvertTo-CodePointList $Pattern
    $script:PI = 0
    try {
        $ast = Read-PAlt
        if ($script:PI -ne $script:PP.Count) { Stop-Pattern 'unexpected' }
    } catch {
        if ($_.Exception.Data.Contains('PATTERN')) {
            Add-Warn ('pattern を解釈できないためランダム英数字で代用しました: ' + $Pattern)
            return $null
        }
        throw
    }
    $rx = $null
    try { $rx = New-Object System.Text.RegularExpressions.Regex($Pattern) } catch { $rx = $null }
    $last = $null
    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        $sb = New-Object System.Text.StringBuilder
        Write-PNode $ast $sb
        $s = $sb.ToString()
        $last = $s
        if (($null -ne $rx) -and -not $rx.IsMatch($s)) { continue }
        if (($null -ne $minL) -and ((Get-CpLen $s) -lt $minL)) { continue }
        if (($null -ne $maxL) -and ((Get-CpLen $s) -gt $maxL)) { continue }
        return $s
    }
    if (($null -ne $last) -and (($null -eq $rx) -or $rx.IsMatch($last))) { return $last }
    Add-Warn ('pattern に一致する値を生成できないためランダム英数字で代用しました: ' + $Pattern)
    return $null
}

# ----------------------------------------------------------------- パラメータ
$script:METHOD_KEYS = @('get', 'put', 'post', 'delete', 'options', 'head', 'patch', 'trace')
$script:ANY_KEY = 'x-amazon-apigateway-any-method'
$script:SKIP_HEADERS = @('accept', 'content-type', 'authorization')
$script:MANAGED_HEADERS = @('content-length', 'host', 'transfer-encoding', 'connection')

function Get-Params($item, $op) {
    $table = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($src in @((Get-OaProp $item 'parameters'), (Get-OaProp $op 'parameters'))) {
        if (-not (Test-Arr $src)) { continue }
        foreach ($p in $src) {
            $rp = Resolve-Deref $p 'parameter'
            if (-not (Test-Obj $rp)) { continue }
            $name = Get-OaProp $rp 'name'
            $loc = Get-OaProp $rp 'in'
            if (-not (Test-Str $name) -or -not (Test-Str $loc) -or ($name.Length -eq 0)) { continue }
            $table[$loc + "`n" + $name] = $rp
        }
    }
    $list = New-Object 'System.Collections.Generic.List[object]'
    foreach ($v in $table.Values) { $list.Add($v) }
    return , $list
}

function Get-FirstValue($d) {
    foreach ($k in @($d.Keys)) { return , ($d[$k]) }
    return $null
}

function Get-ParamValue($p) {
    if ($script:OptUseExamples) {
        if (Test-Has $p 'example') { return , (ConvertTo-GenValue $p['example']) }
        $exs = Get-OaProp $p 'examples'
        if ((Test-Obj $exs) -and ($exs.Count -gt 0)) {
            $first = Resolve-Deref (Get-FirstValue $exs) 'example'
            if ((Test-Obj $first) -and (Test-Has $first 'value')) { return , (ConvertTo-GenValue $first['value']) }
        }
    }
    $schema = Get-OaProp $p 'schema'
    if ($null -eq $schema) {
        $content = Get-OaProp $p 'content'
        if ((Test-Obj $content) -and ($content.Count -gt 0)) { $schema = Get-OaProp (Get-FirstValue $content) 'schema' }
    }
    if ($null -eq $schema) {
        $schema = New-OMap
        $schema['type'] = 'string'
    }
    $empty = New-Object 'System.Collections.Generic.List[string]'
    $v = New-Gen $schema 0 $empty
    if (Test-Skip $v) { $v = '' }
    return , $v
}

function Get-SimpleStyle($v, [bool]$Explode) {
    $parts = New-Object 'System.Collections.Generic.List[string]'
    if (Test-Arr $v) {
        foreach ($x in $v) { $parts.Add((Get-ScalarStr $x)) }
        return [string]::Join(',', $parts)
    }
    if (Test-Obj $v) {
        foreach ($k in @($v.Keys)) {
            if ($Explode) { $parts.Add([string]$k + '=' + (Get-ScalarStr $v[$k])) }
            else { $parts.Add([string]$k + ',' + (Get-ScalarStr $v[$k])) }
        }
        return [string]::Join(',', $parts)
    }
    return (Get-ScalarStr $v)
}

function New-Pair([string]$Name, [string]$Value) { return , @($Name, $Value) }

function Get-QueryPairs([string]$Name, $v, [string]$Style, [bool]$Explode) {
    $pairs = New-Object 'System.Collections.Generic.List[object]'
    if (Test-Arr $v) {
        if (($Style -ceq 'form') -and $Explode) {
            foreach ($x in $v) { $pairs.Add((New-Pair $Name (Get-ScalarStr $x))) }
            return , $pairs
        }
        if ($Style -ceq 'spaceDelimited') { $sep = ' ' } elseif ($Style -ceq 'pipeDelimited') { $sep = '|' } else { $sep = ',' }
        $parts = New-Object 'System.Collections.Generic.List[string]'
        foreach ($x in $v) { $parts.Add((Get-ScalarStr $x)) }
        $pairs.Add((New-Pair $Name ([string]::Join($sep, $parts))))
        return , $pairs
    }
    if (Test-Obj $v) {
        if ($Style -ceq 'deepObject') {
            foreach ($k in @($v.Keys)) { $pairs.Add((New-Pair ($Name + '[' + [string]$k + ']') (Get-ScalarStr $v[$k]))) }
            return , $pairs
        }
        if ($Explode) {
            foreach ($k in @($v.Keys)) { $pairs.Add((New-Pair ([string]$k) (Get-ScalarStr $v[$k]))) }
            return , $pairs
        }
        $parts = New-Object 'System.Collections.Generic.List[string]'
        foreach ($k in @($v.Keys)) { $parts.Add([string]$k + ',' + (Get-ScalarStr $v[$k])) }
        $pairs.Add((New-Pair $Name ([string]::Join(',', $parts))))
        return , $pairs
    }
    $pairs.Add((New-Pair $Name (Get-ScalarStr $v)))
    return , $pairs
}

# ----------------------------------------------------------------- リクエストボディ
function Get-MediaBase([string]$mt) { return $mt.Split([char]';', 2)[0].Trim().ToLowerInvariant() }

function Select-Media($keys) {
    # JSON 系 → フォーム → マルチパート の順に優先し、どれも無ければ定義の先頭のメディアタイプを使う
    for ($pi = 0; $pi -lt 5; $pi++) {
        foreach ($k in $keys) {
            $b = Get-MediaBase $k
            $hit = $false
            switch ($pi) {
                0 { $hit = ($b -ceq 'application/json') }
                1 { $hit = $b.EndsWith('+json', [System.StringComparison]::Ordinal) }
                2 { $hit = $b.Contains('json') }
                3 { $hit = ($b -ceq 'application/x-www-form-urlencoded') }
                4 { $hit = ($b -ceq 'multipart/form-data') }
            }
            if ($hit) { return $k }
        }
    }
    return $keys[0]
}

function Get-MediaKind([string]$mt) {
    $b = Get-MediaBase $mt
    if (($b -ceq 'application/json') -or $b.EndsWith('+json', [System.StringComparison]::Ordinal) -or $b.Contains('json') -or ($b -ceq '*/*') -or ($b -ceq 'application/*')) { return 'json' }
    if ($b -ceq 'application/x-www-form-urlencoded') { return 'form' }
    if ($b.StartsWith('multipart/', [System.StringComparison]::Ordinal)) { return 'multipart' }
    if ($b.EndsWith('/xml', [System.StringComparison]::Ordinal) -or $b.EndsWith('+xml', [System.StringComparison]::Ordinal)) { return 'xml' }
    if ($b.StartsWith('text/', [System.StringComparison]::Ordinal)) { return 'text' }
    return 'binary'
}

function Get-ContentTypeFor([string]$mt, [string]$Kind) {
    if ((Get-MediaBase $mt).Contains('*')) {
        switch -CaseSensitive ($Kind) {
            'json' { return 'application/json' }
            'xml' { return 'application/xml' }
            'text' { return 'text/plain' }
        }
        return 'application/octet-stream'
    }
    return $mt.Trim()
}

function Get-XmlName([string]$n) {
    $n = [regex]::Replace($n, '[^A-Za-z0-9_.-]', '_')
    if (($n.Length -eq 0) -or -not ($n -cmatch '^[A-Za-z_]')) { $n = '_' + $n }
    return $n
}

function Get-XmlText([string]$s) {
    return (Get-CleanText $s).Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function ConvertTo-XmlText($v, [string]$Name, [int]$Indent) {
    $pad = '  ' * $Indent
    if (Test-Obj $v) {
        if ($v.Count -eq 0) { return $pad + '<' + $Name + '/>' }
        $lines = New-Object 'System.Collections.Generic.List[string]'
        $lines.Add($pad + '<' + $Name + '>')
        foreach ($k in @($v.Keys)) {
            $kn = Get-XmlName ([string]$k)
            $x = $v[$k]
            if (Test-Arr $x) {
                foreach ($it in $x) { $lines.Add((ConvertTo-XmlText $it $kn ($Indent + 1))) }
            } else {
                $lines.Add((ConvertTo-XmlText $x $kn ($Indent + 1)))
            }
        }
        $lines.Add($pad + '</' + $Name + '>')
        return [string]::Join("`n", $lines)
    }
    if (Test-Arr $v) {
        if ($v.Count -eq 0) { return $pad + '<' + $Name + '/>' }
        $lines = New-Object 'System.Collections.Generic.List[string]'
        foreach ($it in $v) { $lines.Add((ConvertTo-XmlText $it $Name $Indent)) }
        return [string]::Join("`n", $lines)
    }
    return $pad + '<' + $Name + '>' + (Get-XmlText (Get-ScalarStr $v)) + '</' + $Name + '>'
}

function Get-XmlRootName($schema) {
    $s = $schema
    $name = $null
    $guard = 0
    while ((Test-Obj $s) -and (Test-Has $s '$ref') -and ($guard -lt 32)) {
        $ref = $s['$ref']
        if (Test-Str $ref) {
            $segs = $ref.Split([char]'/')
            $name = $segs[$segs.Length - 1]
        }
        $s = Resolve-Ref $ref
        $guard++
    }
    $x = Get-OaProp $s 'xml'
    if ((Test-Obj $x) -and (Test-Str (Get-OaProp $x 'name')) -and ($x['name'].Length -gt 0)) { $name = $x['name'] }
    if ([string]::IsNullOrEmpty($name)) { $name = 'root' }
    return (Get-XmlName $name)
}

function New-Body($op) {
    if (-not (Test-Has $op 'requestBody')) { return $null }
    $rb = Resolve-Deref $op['requestBody'] 'requestBody'
    $content = Get-OaProp $rb 'content'
    if (-not (Test-Obj $content) -or ($content.Count -eq 0)) { return $null }
    $keys = New-Object 'System.Collections.Generic.List[string]'
    foreach ($k in @($content.Keys)) { $keys.Add([string]$k) }
    $mt = Select-Media $keys
    $media = $content[$mt]
    $schema = Get-OaProp $media 'schema'
    $kind = Get-MediaKind $mt
    $have = $false
    $value = $null
    if ($script:OptUseExamples -and (Test-Obj $media)) {
        if (Test-Has $media 'example') {
            $value = ConvertTo-GenValue $media['example']
            $have = $true
        } else {
            $exs = Get-OaProp $media 'examples'
            if ((Test-Obj $exs) -and ($exs.Count -gt 0)) {
                $first = Resolve-Deref (Get-FirstValue $exs) 'example'
                if ((Test-Obj $first) -and (Test-Has $first 'value')) {
                    $value = ConvertTo-GenValue $first['value']
                    $have = $true
                }
            }
        }
    }
    if (-not $have) {
        if ($null -eq $schema) {
            if (@('json', 'form', 'multipart', 'xml') -ccontains $kind) { $value = New-OMap } else { $value = Get-RandStr $script:ALNUM 16 }
        } else {
            $empty = New-Object 'System.Collections.Generic.List[string]'
            $value = New-Gen $schema 0 $empty
            if (Test-Skip $value) { $value = New-OMap }
        }
    }
    $ctype = Get-ContentTypeFor $mt $kind
    $pairs = New-Object 'System.Collections.Generic.List[object]'
    if ($kind -ceq 'json') {
        return @{ Kind = $kind; CType = $ctype; Text = (ConvertTo-JsonText $value 0 $true); Pairs = $pairs }
    }
    if (($kind -ceq 'form') -or ($kind -ceq 'multipart')) {
        if (Test-Obj $value) {
            foreach ($k in @($value.Keys)) {
                $x = $value[$k]
                if (Test-Arr $x) {
                    foreach ($it in $x) { $pairs.Add((New-Pair ([string]$k) (Get-ScalarStr $it))) }
                } else {
                    $pairs.Add((New-Pair ([string]$k) (Get-ScalarStr $x)))
                }
            }
        } else {
            $pairs.Add((New-Pair 'value' (Get-ScalarStr $value)))
        }
        return @{ Kind = $kind; CType = $ctype; Text = ''; Pairs = $pairs }
    }
    if ($kind -ceq 'xml') {
        $text = '<?xml version="1.0" encoding="UTF-8"?>' + "`n" + (ConvertTo-XmlText $value (Get-XmlRootName $schema) 0)
        return @{ Kind = $kind; CType = $ctype; Text = $text; Pairs = $pairs }
    }
    if ($value -is [string]) { $text = $value } else { $text = ConvertTo-JsonText $value 0 $true }
    return @{ Kind = $kind; CType = $ctype; Text = $text; Pairs = $pairs }
}

# ----------------------------------------------------------------- レスポンス（Accept・期待ステータス）
function Test-Is2xx([string]$k) { return (($k -cmatch '^2[0-9][0-9]$') -or ($k -ceq '2XX')) }

function Get-Accept($op) {
    $responses = Get-OaProp $op 'responses'
    if (-not (Test-Obj $responses)) { return $null }
    foreach ($k in @($responses.Keys)) {
        if (-not (Test-Is2xx (([string]$k).Trim().ToUpperInvariant()))) { continue }
        $r = Resolve-Deref $responses[$k] 'response'
        $content = Get-OaProp $r 'content'
        if ((Test-Obj $content) -and ($content.Count -gt 0)) {
            $mt = [string](@($content.Keys)[0])
            if ($mt.Contains('*')) { return $null }
            return $mt.Trim()
        }
    }
    return $null
}

function Get-SuccessCodes($op) {
    $responses = Get-OaProp $op 'responses'
    $codes = New-Object 'System.Collections.Generic.List[string]'
    $hasRange = $false
    if (Test-Obj $responses) {
        foreach ($k in @($responses.Keys)) {
            $ks = ([string]$k).Trim().ToUpperInvariant()
            if ($ks -cmatch '^2[0-9][0-9]$') {
                if (-not $codes.Contains($ks)) { $codes.Add($ks) }
            } elseif ($ks -ceq '2XX') {
                $hasRange = $true
            }
        }
    }
    return @{ Codes = $codes; HasRange = $hasRange }
}

# ----------------------------------------------------------------- セキュリティ
function Get-EffectiveSecurity($op) {
    if (Test-Has $op 'security') { return , ($op['security']) }
    return , (Get-OaProp $script:Doc 'security')
}

function Get-SecurityText($op) {
    $sec = Get-EffectiveSecurity $op
    if (-not (Test-Arr $sec)) { return $null }
    if ($sec.Count -eq 0) { return 'なし（認証不要）' }
    $alts = New-Object 'System.Collections.Generic.List[string]'
    foreach ($req in $sec) {
        if (-not (Test-Obj $req)) { continue }
        if ($req.Count -eq 0) { $alts.Add('認証なしも可'); continue }
        $parts = New-Object 'System.Collections.Generic.List[string]'
        foreach ($nk in @($req.Keys)) {
            $scopes = $req[$nk]
            $sc = New-Object 'System.Collections.Generic.List[string]'
            if (Test-Arr $scopes) { foreach ($x in $scopes) { if (Test-Str $x) { $sc.Add($x) } } }
            if ($sc.Count -gt 0) { $parts.Add([string]$nk + '[' + [string]::Join(', ', $sc) + ']') } else { $parts.Add([string]$nk) }
        }
        $alts.Add([string]::Join(' かつ ', $parts))
    }
    if ($alts.Count -eq 0) { return $null }
    return [string]::Join(' または ', $alts)
}

function Test-SchemeUsesAuthorization($scheme) {
    if (-not (Test-Obj $scheme)) { return $false }
    $t = Get-OaProp $scheme 'type'
    if (-not (Test-Str $t)) { return $false }
    if (($t -ceq 'oauth2') -or ($t -ceq 'openIdConnect') -or ($t -ceq 'http')) { return $true }
    if ($t -ceq 'apiKey') {
        $in = Get-OaProp $scheme 'in'
        $nm = Get-OaProp $scheme 'name'
        return ((Test-Str $in) -and ($in -ceq 'header') -and (Test-Str $nm) -and ($nm.ToLowerInvariant() -ceq 'authorization'))
    }
    return $false
}

function Test-NeedsAuthHeader($ops) {
    $schemes = Get-OaProp (Get-OaProp $script:Doc 'components') 'securitySchemes'
    foreach ($info in $ops) {
        $sec = Get-EffectiveSecurity $info.Op
        if (-not (Test-Arr $sec)) { continue }
        foreach ($req in $sec) {
            if (-not (Test-Obj $req)) { continue }
            foreach ($nk in @($req.Keys)) {
                $sch = Resolve-Deref (Get-OaProp $schemes ([string]$nk)) 'securityScheme'
                if (Test-SchemeUsesAuthorization $sch) { return $true }
            }
        }
    }
    return $false
}

# ----------------------------------------------------------------- オペレーション
function Get-Operations {
    $ops = New-Object 'System.Collections.Generic.List[object]'
    $paths = Get-OaProp $script:Doc 'paths'
    foreach ($pk in @($paths.Keys)) {
        $path = [string]$pk
        if ($path.StartsWith('x-', [System.StringComparison]::Ordinal)) { continue }
        $item = Resolve-Deref $paths[$pk] 'pathItem'
        if (-not (Test-Obj $item)) { continue }
        foreach ($kk in @($item.Keys)) {
            $key = [string]$kk
            $k = $key.ToLowerInvariant()
            if ($script:METHOD_KEYS -ccontains $k) {
                $method = $k.ToUpperInvariant()
                $isAny = $false
            } elseif ($k -ceq $script:ANY_KEY) {
                $method = $script:OptAnyMethod
                $isAny = $true
            } else {
                continue
            }
            $op = $item[$kk]
            if (-not (Test-Obj $op)) { continue }
            $ops.Add(@{ Path = $path; Item = $item; Op = $op; Key = $key; Method = $method; IsAny = $isAny })
        }
    }
    return , $ops
}

$script:PLACEHOLDER_RE = New-Object System.Text.RegularExpressions.Regex('\{([^{}]+)\}')

function New-OperationRequest($info, [string]$Label) {
    $path = $info.Path
    $op = $info.Op
    $method = $info.Method
    $params = Get-Params $info.Item $op

    # 1. パスパラメータ
    $pathVals = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($p in $params) {
        if ((Get-OaProp $p 'in') -ceq 'path') {
            $ex = Get-OaProp $p 'explode'
            $pathVals[[string]$p['name']] = @{ V = (Get-ParamValue $p); Explode = (($ex -is [bool]) -and $ex) }
        }
    }
    foreach ($m in $script:PLACEHOLDER_RE.Matches($path)) {
        $ph = $m.Groups[1].Value
        if ($ph.EndsWith('+', [System.StringComparison]::Ordinal)) { $base = $ph.Substring(0, $ph.Length - 1) } else { $base = $ph }
        if (-not $pathVals.Contains($base)) {
            Add-Warn ('{0}: パス変数 {{{1}}} のパラメータ定義が無いため英数字 8 文字を仮設定しました' -f $Label, $base)
            $pathVals[$base] = @{ V = (Get-RandStr $script:ALNUM 8); Explode = $false }
        }
    }
    $sb = New-Object System.Text.StringBuilder
    $pos = 0
    foreach ($m in $script:PLACEHOLDER_RE.Matches($path)) {
        [void]$sb.Append($path.Substring($pos, $m.Index - $pos))
        $ph = $m.Groups[1].Value
        $greedy = $ph.EndsWith('+', [System.StringComparison]::Ordinal)
        if ($greedy) { $base = $ph.Substring(0, $ph.Length - 1) } else { $base = $ph }
        $pv = $pathVals[$base]
        [void]$sb.Append((Get-PctEncode (Get-SimpleStyle $pv.V $pv.Explode) $greedy))
        $pos = $m.Index + $m.Length
    }
    [void]$sb.Append($path.Substring($pos))
    $realPath = $sb.ToString()
    if (-not $realPath.StartsWith('/', [System.StringComparison]::Ordinal)) { $realPath = '/' + $realPath }

    # 2. クエリパラメータ
    $qpairs = New-Object 'System.Collections.Generic.List[object]'
    foreach ($p in $params) {
        if ((Get-OaProp $p 'in') -cne 'query') { continue }
        $rq = Get-OaProp $p 'required'
        if ($script:OptRequiredOnly -and -not (($rq -is [bool]) -and $rq)) { continue }
        $v = Get-ParamValue $p
        $style = Get-OaProp $p 'style'
        if (-not (Test-Str $style)) { $style = 'form' }
        $explode = Get-OaProp $p 'explode'
        if (-not ($explode -is [bool])) { $explode = ($style -ceq 'form') }
        foreach ($pair in (Get-QueryPairs ([string]$p['name']) $v $style $explode)) { $qpairs.Add($pair) }
    }

    # 3. ヘッダパラメータ
    $hparams = New-Object 'System.Collections.Generic.List[object]'
    foreach ($p in $params) {
        if ((Get-OaProp $p 'in') -cne 'header') { continue }
        $name = [string]$p['name']
        $lname = $name.ToLowerInvariant()
        if ($script:SKIP_HEADERS -ccontains $lname) { continue }
        if ($lname -ceq 'api-key') {
            Add-Warn ('{0}: ヘッダ {1} は API キーの実値が必要なため生成しません（OpenAPI 用設定ファイル（-OpenApiConfig）の API_HEADERS に「{1}: キーの値」を指定してください）' -f $Label, $name)
            continue
        }
        if ($script:MANAGED_HEADERS -ccontains $lname) {
            Add-Warn ('{0}: ヘッダ {1} は HTTP クライアントが自動設定するため生成しません' -f $Label, $name)
            continue
        }
        $rq = Get-OaProp $p 'required'
        if ($script:OptRequiredOnly -and -not (($rq -is [bool]) -and $rq)) { continue }
        $v = Get-ParamValue $p
        $ex = Get-OaProp $p 'explode'
        $hparams.Add((New-Pair $name (Get-HeaderSafe (Get-SimpleStyle $v (($ex -is [bool]) -and $ex)))))
    }

    # 4. クッキーパラメータ
    $cookies = New-Object 'System.Collections.Generic.List[string]'
    foreach ($p in $params) {
        if ((Get-OaProp $p 'in') -cne 'cookie') { continue }
        $rq = Get-OaProp $p 'required'
        if ($script:OptRequiredOnly -and -not (($rq -is [bool]) -and $rq)) { continue }
        $v = Get-ParamValue $p
        $cookies.Add([string]$p['name'] + '=' + (Get-PctEncode (Get-SimpleStyle $v $false) $false))
    }

    # 5. リクエストボディ
    $body = New-Body $op

    # 6. 組み立て（API一覧（Excel / CSV）の 1 行と同じ項目を返す。クエリを Parameters に入れるか URL に
    #    付けるかは、Excel / CSV の行と同じく Normalize-Api がメソッドで決める）
    $bodyType = ''
    $bodyText = ''
    $headers = New-Object 'System.Collections.Generic.List[object]'
    if ($null -ne $body) {
        if (($body.Kind -ceq 'form') -or ($body.Kind -ceq 'multipart')) {
            $bodyType = $body.Kind
            $bodyText = Get-OaRowPairs $body.Pairs
            if ($body.Kind -ceq 'form') { $headers.Add((New-Pair 'Content-Type' $body.CType)) }
        } else {
            if ($body.Kind -ceq 'json') { $bodyType = 'json' } else { $bodyType = 'raw' }
            $bodyText = $body.Text
            $headers.Add((New-Pair 'Content-Type' $body.CType))
        }
    }
    $accept = Get-Accept $op
    if ($null -ne $accept) { $headers.Add((New-Pair 'Accept' $accept)) }
    foreach ($h in $hparams) { $headers.Add($h) }
    if ($cookies.Count -gt 0) { $headers.Add((New-Pair 'Cookie' ([string]::Join('; ', $cookies)))) }

    $sc = Get-SuccessCodes $op
    $hasData = ($pathVals.Count -gt 0) -or ($qpairs.Count -gt 0) -or ($hparams.Count -gt 0) -or ($cookies.Count -gt 0) -or ($null -ne $body)
    $comment = Get-OperationComment $info $body $hasData
    if ($sc.HasRange -and ($sc.Codes.Count -eq 0)) {
        $comment += "`n" + '※ 成功レスポンスが 2XX（範囲指定）のため期待ステータスは空欄です（JMeter の既定の判定: 2xx / 3xx = 成功）。'
    }
    $lines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($h in $headers) { $lines.Add([string]$h[0] + ': ' + [string]$h[1]) }
    return @{
        method = $method; url = $realPath; query = (Get-OaRowPairs $qpairs); headers = [string]::Join("`n", $lines)
        body_type = $bodyType; body = $bodyText; expected_status = [string]::Join(',', $sc.Codes); comment = $comment
    }
}

function Get-OperationComment($info, $body, [bool]$HasData) {
    $op = $info.Op
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $summary = Get-OaProp $op 'summary'
    if ((Test-Str $summary) -and ($summary.Trim().Length -gt 0)) { $lines.Add('概要: ' + $summary.Trim()) }
    $desc = Get-OaProp $op 'description'
    if ((Test-Str $desc) -and ($desc.Trim().Length -gt 0)) { $lines.Add('説明: ' + $desc.Trim()) }
    $lines.Add('OpenAPI 定義: ' + $info.Key + ' ' + $info.Path)
    $opId = Get-OaProp $op 'operationId'
    if ((Test-Str $opId) -and ($opId.Length -gt 0)) { $lines.Add('operationId: ' + $opId) }
    if ($info.IsAny) {
        $lines.Add('ANY メソッド（x-amazon-apigateway-any-method）のため ' + $info.Method + ' で代表して送信します（生成時の --any-method / -AnyMethod で変更可）。')
    }
    $sec = Get-SecurityText $op
    if ($null -ne $sec) { $lines.Add('認可（security）: ' + $sec) }
    $integ = Get-OaProp $op 'x-amazon-apigateway-integration'
    if (Test-Obj $integ) {
        $items = New-Object 'System.Collections.Generic.List[string]'
        foreach ($k in @('type', 'httpMethod', 'uri', 'connectionType')) {
            $iv = Get-OaProp $integ $k
            if (Test-Str $iv) { $items.Add($k + '=' + $iv) }
        }
        if ($items.Count -gt 0) { $lines.Add('API Gateway 統合: ' + [string]::Join(', ', $items)) }
    }
    $dep = Get-OaProp $op 'deprecated'
    if (($dep -is [bool]) -and $dep) { $lines.Add('※ deprecated（非推奨）に指定された API です。') }
    if (($null -ne $body) -and ($body.Kind -ceq 'multipart')) {
        $lines.Add('※ multipart/form-data はファイル項目もテキスト値で送信します。実ファイルを送る場合は「ファイルアップロード」タブで設定してください。')
    }
    if ($HasData) {
        $lines.Add('テストデータ: OpenAPI のデータ型定義からランダム生成（乱数シード ' + $script:SeedValue.ToString($script:Inv) + '）')
    }
    return [string]::Join("`n", $lines)
}

# ----------------------------------------------------------------- JSON 読み込み
function ConvertFrom-JssValue($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [System.Collections.Generic.Dictionary[string, object]]) {
        $d = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::Ordinal)
        foreach ($k in @($v.Keys)) { $d[$k] = (ConvertFrom-JssValue $v[$k]) }
        return , $d
    }
    if (($v -is [object[]]) -or ($v -is [System.Collections.ArrayList])) {
        $l = New-Object 'System.Collections.Generic.List[object]'
        foreach ($i in $v) { $l.Add((ConvertFrom-JssValue $i)) }
        return , $l
    }
    if ($v -is [int]) { return [long]$v }
    if ($v -is [double]) {
        return [decimal]::Parse($v.ToString('R', $script:Inv), [System.Globalization.NumberStyles]::Float, $script:Inv)
    }
    return $v
}

function ConvertFrom-JsonElementValue($el) {
    switch ([string]$el.ValueKind) {
        'Object' {
            $d = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::Ordinal)
            foreach ($p in $el.EnumerateObject()) { $d[$p.Name] = (ConvertFrom-JsonElementValue $p.Value) }
            return , $d
        }
        'Array' {
            $l = New-Object 'System.Collections.Generic.List[object]'
            foreach ($i in $el.EnumerateArray()) { $l.Add((ConvertFrom-JsonElementValue $i)) }
            return , $l
        }
        'String' { return $el.GetString() }
        'Number' {
            $raw = $el.GetRawText()
            if ($raw -cmatch '^-?[0-9]+$') {
                $lv = 0L
                if ([long]::TryParse($raw, [System.Globalization.NumberStyles]::AllowLeadingSign, $script:Inv, [ref]$lv)) { return $lv }
            }
            return [decimal]::Parse($raw, [System.Globalization.NumberStyles]::Float, $script:Inv)
        }
        'True' { return $true }
        'False' { return $false }
    }
    return $null
}

# JSON 文字列（エスケープ込み）と ${...} を先頭から順に拾う。文字列は読み飛ばし、${...} だけを置き換える。
# 直後が : の ${...} はキーの位置なので置き換えない（5.1 の JavaScriptSerializer は引用符なしのキーを
# 受け付けてしまい、実装間で結果が変わるため。置き換えなければどの実装でも構文エラーになる）
$script:PlaceholderScan = New-Object System.Text.RegularExpressions.Regex('"[^"\\]*(?:\\.[^"\\]*)*"|\$\{[^{}"\r\n]+\}(?![ \t\r\n]*:)', [System.Text.RegularExpressions.RegexOptions]::Singleline)
$script:PlaceholderLinesShown = 10
$script:PlaceholderCount = 0

function Resolve-Placeholders([string]$Text, [string]$Value) {
    # JSON 文字列の外に引用符なしで書かれた ${...}（後で置換するプレースホルダ）を数値 $Value に置き換える。
    # 文字列の中の ${...}（"${stageVariables.x}" など）は正しい JSON なので変更しない。
    # 戻り値: @{ Text = 置換後の文字列; Names = プレースホルダ（初出順）; Lines = プレースホルダ → 行番号の一覧 }
    $names = New-Object 'System.Collections.Generic.List[string]'
    $lines = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::Ordinal)
    if ($Text.IndexOf('${', [System.StringComparison]::Ordinal) -lt 0) { return @{ Text = $Text; Names = $names; Lines = $lines } }
    $sb = New-Object System.Text.StringBuilder
    $pos = 0
    $line = 1L
    foreach ($m in $script:PlaceholderScan.Matches($Text)) {
        $s = $m.Index
        if ($Text[$s] -cne [char]'$') { continue }    # 文字列（読み飛ばすだけ）
        $j = $Text.IndexOf([char]10, $pos)
        while (($j -ge 0) -and ($j -lt $s)) { $line++; $j = $Text.IndexOf([char]10, $j + 1) }
        [void]$sb.Append($Text, $pos, $s - $pos)
        [void]$sb.Append($Value)
        if (-not $lines.ContainsKey($m.Value)) {
            $names.Add($m.Value)
            $lines[$m.Value] = New-Object 'System.Collections.Generic.List[long]'
        }
        $lines[$m.Value].Add($line)
        $pos = $s + $m.Length
    }
    [void]$sb.Append($Text, $pos, $Text.Length - $pos)
    return @{ Text = $sb.ToString(); Names = $names; Lines = $lines }
}

function Get-PlaceholderWarning([string]$Name, $Lines, [string]$Value) {
    $shown = New-Object 'System.Collections.Generic.List[string]'
    for ($i = 0; ($i -lt $Lines.Count) -and ($i -lt $script:PlaceholderLinesShown); $i++) { $shown.Add($Lines[$i].ToString($script:Inv)) }
    if ($Lines.Count -gt $script:PlaceholderLinesShown) { $more = ' ほか' } else { $more = '' }
    return ('引用符なしのプレースホルダ ' + $Name + ' は JSON として読めないため、数値 ' + $Value + ' に置き換えて読み込みました（' +
        $Lines.Count.ToString($script:Inv) + ' 箇所・行 ' + [string]::Join(', ', $shown) + $more + '）')
}

function Read-JsonFile([string]$Path, [string]$PlaceholderValue) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $strict = New-Object System.Text.UTF8Encoding($false, $true)
    try {
        if (($bytes.Length -ge 3) -and ($bytes[0] -eq 0xEF) -and ($bytes[1] -eq 0xBB) -and ($bytes[2] -eq 0xBF)) {
            $text = $strict.GetString($bytes, 3, $bytes.Length - 3)
        } elseif (($bytes.Length -ge 2) -and ($bytes[0] -eq 0xFF) -and ($bytes[1] -eq 0xFE)) {
            $text = [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
        } elseif (($bytes.Length -ge 2) -and ($bytes[0] -eq 0xFE) -and ($bytes[1] -eq 0xFF)) {
            $text = [System.Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2)
        } else {
            $text = $strict.GetString($bytes)
        }
    } catch {
        Stop-Gen ('入力ファイルの文字コードを UTF-8 として読めません: ' + $Path)
    }
    $ph = Resolve-Placeholders $text $PlaceholderValue
    $text = $ph.Text
    try {
        if ($PSVersionTable.PSEdition -eq 'Core') {
            $opts = New-Object System.Text.Json.JsonDocumentOptions
            $opts.MaxDepth = 1024
            $jd = [System.Text.Json.JsonDocument]::Parse($text, $opts)
            try { $doc = ConvertFrom-JsonElementValue $jd.RootElement } finally { $jd.Dispose() }
        } else {
            Add-Type -AssemblyName System.Web.Extensions
            $jss = New-Object System.Web.Script.Serialization.JavaScriptSerializer
            $jss.MaxJsonLength = [int]::MaxValue
            $jss.RecursionLimit = 1024
            $doc = ConvertFrom-JssValue ($jss.DeserializeObject($text))
        }
    } catch {
        if ($_.Exception.Data.Contains('OA2J')) { throw }
        Stop-Gen ('JSON の構文エラーです（YAML 形式は未対応です）: ' + $_.Exception.Message)
    }
    foreach ($nm in $ph.Names) { Add-Warn (Get-PlaceholderWarning $nm $ph.Lines[$nm] $PlaceholderValue) }
    $script:PlaceholderCount = $ph.Names.Count
    return , $doc
}

function Get-OaRowPairs($pairs) {
    # (名前, 値) の並びを API一覧の「名前=値」（改行区切り）の形にする。
    # 区切り文字（& 改行 =）・%・$・前後の空白を含むものは、名前・値とも URL エンコードして書く
    $out = New-Object 'System.Collections.Generic.List[string]'
    foreach ($pr in $pairs) {
        $k = [string]$pr[0]
        $v = [string]$pr[1]
        if (($k -ceq $k.Trim()) -and ($v -ceq $v.Trim()) -and -not $k.Contains('=') -and -not [regex]::IsMatch($k + $v, '[&%$\x00-\x1F\x7F]')) {
            $out.Add($k + '=' + $v)
            continue
        }
        $ek = Get-PctEncode $k $false
        $ev = Get-PctEncode $v $false
        if (-not $ev.Contains('%')) {
            # 値に %XX が無いと JMeter がエンコード済みの名前をもう一度エンコードするため、値もすべて %XX にする
            $sb = New-Object System.Text.StringBuilder
            foreach ($b in [System.Text.Encoding]::UTF8.GetBytes($ev)) { [void]$sb.Append('%').Append(([int]$b).ToString('X2', $script:Inv)) }
            $ev = $sb.ToString()
        }
        $out.Add($ek + '=' + $ev)
    }
    return [string]::Join("`n", $out)
}

$script:OaVars = @(@('public_base_url', 'PUBLIC_BASE_URL'), @('name_prefix', 'NAME_PREFIX'))

function Resolve-OaVars([string]$text, [string]$where) {
    # OpenAPI 定義のプレースホルダ ${public_base_url} / ${name_prefix} を OpenAPI 用設定（-OpenApiConfig）の値に置き換える。
    # 未設定（空欄）の場合は隣の区切り文字（- _ . / 空白）ごと取り除き、パスの先頭やテストプラン名に ${...} を残さない
    foreach ($v in $script:OaVars) {
        $ph = '${' + $v[0] + '}'
        if (-not $text.Contains($ph)) { continue }
        $value = ''
        if ($null -ne $script:OaConf) { $value = Get-OaConf $script:OaConf $v[1] }
        if ($value) {
            $text = $text.Replace($ph, $value)
            Write-Info ('OpenAPI 定義 ' + $where + ' の ' + $ph + ': ' + $value + '（OpenAPI 用設定 ' + $v[1] + ' の値）')
        } else {
            $text = $text -creplace ('^\$\{' + $v[0] + '\}[-_. /]*|[-_. /]*\$\{' + $v[0] + '\}'), ''
            Write-Info ('OpenAPI 定義 ' + $where + ' の ' + $ph + ': 未設定のため取り除きます（OpenAPI 用設定ファイル（-OpenApiConfig）の ' + $v[1] + ' で設定できます）')
        }
    }
    return $text
}

function Get-OaServerUrl($doc) {
    # 接続先の既定値: servers[0].url（${public_base_url} は OpenAPI 用設定の値、{変数} は既定値に置換）。http(s) の URL でなければ http://localhost:8080 を前に付ける
    $servers = Get-OaProp $doc 'servers'
    $srv = $null
    if ((Test-Arr $servers) -and ($servers.Count -gt 0) -and (Test-Obj $servers[0])) { $srv = $servers[0] }
    $url = Get-OaProp $srv 'url'
    if (Test-Str $url) { $url = $url.Trim() } else { $url = '' }
    $url = Resolve-OaVars $url 'servers[0].url'
    $vars = Get-OaProp $srv 'variables'
    $sb = New-Object System.Text.StringBuilder
    $pos = 0
    foreach ($m in $script:PLACEHOLDER_RE.Matches($url)) {
        [void]$sb.Append($url.Substring($pos, $m.Index - $pos))
        $d = Get-OaProp (Get-OaProp $vars $m.Groups[1].Value) 'default'
        if (Test-Str $d) { [void]$sb.Append($d) } else { [void]$sb.Append($m.Value) }
        $pos = $m.Index + $m.Length
    }
    [void]$sb.Append($url.Substring($pos))
    $url = $sb.ToString().TrimEnd([char]'/')
    if ($url -cmatch '^[Hh][Tt][Tt][Pp][Ss]?://') { return $url }
    if (($url.Length -gt 0) -and -not $url.StartsWith('/', [System.StringComparison]::Ordinal)) { $url = '/' + $url }
    return 'http://localhost:8080' + $url
}

function Load-OpenApi([string]$inputPath) {
    # OpenAPI 3.0.x 定義（JSON）を読み込み、API一覧（Excel / CSV）と同じ項目の行に変換する
    # 戻り値: @(行の一覧, OpenAPI 定義から求めた設定の既定値 @(@(設定キー, 値), ...))
    if ($Seed.Length -gt 0) {
        if (-not ($Seed -cmatch '^[0-9]{1,12}\z')) { throw (New-ToolError ('-Seed は 0 以上の整数（12 桁以内）で指定してください: ' + $Seed)) }
        $script:SeedValue = [long]::Parse($Seed, $script:Inv)
    } else {
        $buf = New-Object byte[] 4
        $rngc = [System.Security.Cryptography.RandomNumberGenerator]::Create()
        try { $rngc.GetBytes($buf) } finally { $rngc.Dispose() }
        $script:SeedValue = (([long]$buf[0] * 16777216L) + ([long]$buf[1] * 65536L) + ([long]$buf[2] * 256L) + [long]$buf[3]) % 1000000000L
    }
    $anyM = $AnyMethod.ToUpperInvariant()
    if (-not (@('GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'HEAD', 'OPTIONS') -ccontains $anyM)) { throw (New-ToolError ('-AnyMethod は GET/POST/PUT/PATCH/DELETE/HEAD/OPTIONS のいずれかです: ' + $AnyMethod)) }
    if (-not ($PlaceholderValue -cmatch '^-?(0|[1-9][0-9]{0,14})([.][0-9]{1,12})?\z')) { throw (New-ToolError ('-PlaceholderValue は JSON の数値（例: 3000、-1、2.5。整数部 15 桁・小数部 12 桁以内）で指定してください: ' + $PlaceholderValue)) }
    $script:OptRequiredOnly = [bool]$RequiredOnly
    $script:OptUseExamples = [bool]$UseExamples
    $script:OptAnyMethod = $anyM
    $doc = Read-JsonFile (FullPath $inputPath) $PlaceholderValue
    if (-not (Test-Obj $doc)) { Stop-Gen 'OpenAPI 定義のルートが JSON オブジェクトではありません' }
    $script:Doc = $doc
    if (Test-Has $doc 'swagger') { Stop-Gen 'Swagger 2.0 形式には対応していません。OpenAPI 3.0.x 形式（"openapi": "3.0.3" など）に変換してください' }
    $ver = Get-OaProp $doc 'openapi'
    if (-not (Test-Str $ver) -or -not ($ver -cmatch '^3\.[0-9]+(\.[0-9]+)?')) { Stop-Gen '"openapi" フィールドが 3.x ではありません（OpenAPI 3.0.x の JSON を指定してください）' }
    if (-not $ver.StartsWith('3.0', [System.StringComparison]::Ordinal)) { Write-Warn ('OpenAPI ' + $ver + ' は 3.0.x 以外です。3.0.x と共通の項目のみを解釈します。') }
    if (-not (Test-Obj (Get-OaProp $doc 'paths'))) { Stop-Gen '"paths" が定義されていません' }
    $infoObj = Get-OaProp $doc 'info'
    $title = Get-OaProp $infoObj 'title'
    if (Test-Str $title) { $title = Resolve-OaVars $title 'info.title' }
    if ((Test-Str $title) -and ($title.Trim().Length -gt 0)) { $title = $title.Trim() } else { $title = 'OpenAPI' }
    $apiVersion = Get-OaProp $infoObj 'version'
    if ((Test-Str $apiVersion) -and ($apiVersion.Trim().Length -gt 0)) { $apiVersion = $apiVersion.Trim() } else { $apiVersion = '-' }

    Initialize-Rng $script:SeedValue
    $ops = Get-Operations
    if ($ops.Count -eq 0) { Stop-Gen '"paths" に呼び出し可能な操作（get/post/put/delete など）が 1 件もありません' }
    $records = New-Object System.Collections.Generic.List[object]
    for ($idx = 0; $idx -lt $ops.Count; $idx++) {
        $io = $ops[$idx]
        if ($io.IsAny) { $name = 'ANY(' + $io.Method + ')' } else { $name = $io.Method }
        $name = $name + ' ' + $io.Path
        $opId = Get-OaProp $io.Op 'operationId'
        if ((Test-Str $opId) -and ($opId.Length -gt 0)) { $name += ' [' + $opId + ']' }
        $row = New-OperationRequest $io ('API' + ($idx + 1).ToString('00', $script:Inv) + '_' + $name)
        $row['no'] = ($idx + 1).ToString($script:Inv)
        $row['enabled'] = 'Y'
        $row['name'] = $name
        $rec = @{ _row = ($idx + 1) }
        foreach ($c in $script:Spec.api_columns) {
            $v = ''
            if ($row.ContainsKey($c.key)) { $v = [string]$row[$c.key] }
            $rec[$c.key] = (Clean $v).Trim()
        }
        $records.Add($rec)
    }
    if ((Test-NeedsAuthHeader $ops) -and -not $script:OaAuth) {
        Add-Warn '認可（security）が定義された API があります。Authorization ヘッダが必要な場合は、設定 common_headers（例: Authorization: Bearer ${AUTH_TOKEN}）か、OpenAPI 用設定ファイル（-OpenApiConfig）の TOKEN_DIR で指定してください'
    }
    foreach ($w in $script:OaWarnings) { Write-Warn $w }
    Write-Info ('API定義: OpenAPI ' + $ver + '（' + $title + ' ' + $apiVersion + '、' + $records.Count + ' 件）')
    Write-Info ('テストデータの乱数シード: ' + $script:SeedValue + '（-Seed ' + $script:SeedValue + ' を指定すると同じテストデータを再生成できます）')
    $defaults = @(@('base_url', (Get-OaServerUrl $doc)), @('output_file', ((Get-SafeName $title) + '.jmx')),
        @('test_plan_name', ($title + ' ' + $apiVersion + '（OpenAPI ' + $ver + '）')))
    return @($records, $defaults)
}

# ----------------------------------------------------------------------------- OpenAPI 用設定ファイル（-OpenApiConfig）
# OpenAPI 定義（.json）から生成する場合だけ使う設定ファイル（設定 CSV と同じ形式）。設定ファイル・OpenAPI 定義の値より
# 優先して、次の設定に反映する（RHEL 版 jmeter_api_tool.sh の -c と同じ処理）
#   PROTOCOL / HOST / PORT → base_url の接続先（パス以降はそのまま）、HEADERS → common_headers、
#   TOKEN_DIR → access_token_dir / access_token_file（フォルダ直下の *.token のうち更新日時が最も新しいファイル）
#   PUBLIC_BASE_URL / NAME_PREFIX → OpenAPI 定義の ${public_base_url}（servers[0].url）/ ${name_prefix}（info.title）の値（Resolve-OaVars）
#   API_HEADERS → 各 API のヘッダ（API一覧の「ヘッダ」。-OpenApiConfig を指定しない場合・空欄は既定値。Add-OaApiHeaders）
$script:TokenPattern = '*.token'
$script:OaAuth = $false
$script:OaConf = $null

function Get-OaConf($conf, [string]$k) {
    if ($conf.ContainsKey($k)) { return [string]$conf[$k] }
    return ''
}
function Load-OaConfig([string]$path) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw (New-ToolError ('OpenAPI 用設定ファイルが見つかりません: ' + $path)) }
    $tx = Read-TextAuto $path
    $values = Load-SettingsTable (Rows-FromCsv (Parse-Csv $tx[0])) ([System.IO.Path]::GetFileName($path))
    $known = New-Object System.Collections.Generic.List[string]
    foreach ($c in $script:Spec.openapi_config) { $known.Add([string]$c.key) }
    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($k in $values.Keys) { $keys.Add($k) }
    $keys.Sort([System.StringComparer]::Ordinal)
    $conf = New-Object 'System.Collections.Generic.Dictionary[string,string]'
    foreach ($k in $keys) {
        if ($known.Contains($k.ToUpperInvariant())) { $conf[$k.ToUpperInvariant()] = $values[$k] }
        else { Write-Warn ('OpenAPI 用設定: 不明な設定キー「' + $k + '」は無視します') }
    }
    Write-Info ('OpenAPI 用設定: ' + $path)
    return ,$conf
}
function Apply-OaConfig($raw, $conf, [string]$outDir) {
    # 接続先: 指定した項目だけ base_url を置き換える（空欄の PROTOCOL / HOST は OpenAPI 定義・設定ファイルの値のまま、空欄の PORT は未設定）
    $proto = (Get-OaConf $conf 'PROTOCOL').ToLowerInvariant()
    $hostName = Get-OaConf $conf 'HOST'
    $port = Get-OaConf $conf 'PORT'
    if (-not (@('', 'http', 'https') -ccontains $proto)) {
        Write-Warn ('OpenAPI 用設定 PROTOCOL: 「' + (Get-OaConf $conf 'PROTOCOL') + '」は http / https ではないため無視します')
        $proto = ''
    }
    if ($hostName -and -not ($hostName -cmatch '^(\[[0-9A-Fa-f:.]+\]|[A-Za-z0-9._-]+)\z')) {
        Write-Warn ('OpenAPI 用設定 HOST: 「' + $hostName + '」はホスト名に使えない文字を含むため無視します（例: api.example.com）')
        $hostName = ''
    }
    if ($port -and -not (($port -cmatch '^[0-9]{1,5}\z') -and ([int]$port -ge 1) -and ([int]$port -le 65535))) {
        Write-Warn ('OpenAPI 用設定 PORT: 「' + $port + '」は 1～65535 の数値ではないため無視します')
        $port = ''
    }
    if ($proto -or $hostName -or $port) {
        $base = ''
        if ($raw.ContainsKey('base_url')) { $base = [string]$raw['base_url'] }
        if (-not [regex]::IsMatch($base, $script:UrlRe)) { $base = 'https://' + $base }    # Normalize-Api と同じく、スキームが無い場合は https とみなす
        $m = [regex]::Match($base, $script:UrlRe)
        $p = Parse-Url $base
        if ($m.Success -and $null -ne $p) {
            if (-not $proto) { $proto = $p.scheme }
            if (-not $hostName) { $hostName = $p.host }
            $url = $proto + '://' + $hostName
            if ($port) { $url += ':' + $port }
            $raw['base_url'] = $url + $base.Substring($m.Groups[2].Index + $m.Groups[2].Length)
            Write-Info ('設定 base_url: ' + $raw['base_url'] + '（OpenAPI 用設定の PROTOCOL / HOST / PORT を反映）')
        }
    }
    # HTTP ヘッダ: 共通ヘッダに追加し、同じ名前のヘッダは置き換える。TOKEN_DIR があれば Authorization: Bearer ${ACCESS_TOKEN}
    $add = Parse-Headers (Get-OaConf $conf 'HEADERS') 'OpenAPI 用設定 HEADERS'
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($h in $add) { $names.Add(([string]$h[0]).ToLowerInvariant()) }
    if ((Get-OaConf $conf 'TOKEN_DIR') -and -not $names.Contains('authorization')) {
        $add.Add(@('Authorization', 'Bearer ${ACCESS_TOKEN}'))
        $names.Add('authorization')
    }
    if ($add.Count -gt 0) {
        $lines = New-Object System.Collections.Generic.List[string]
        $ch = ''
        if ($raw.ContainsKey('common_headers')) { $ch = [string]$raw['common_headers'] }
        foreach ($x in (NormNl $ch).Split([char]10)) {
            if ($x.Trim() -and -not $names.Contains($x.Split([char]':', 2)[0].Trim().ToLowerInvariant())) { $lines.Add($x) }
        }
        $shown = New-Object System.Collections.Generic.List[string]
        foreach ($h in $add) { $lines.Add([string]$h[0] + ': ' + [string]$h[1]); $shown.Add([string]$h[0]) }
        $raw['common_headers'] = [string]::Join("`n", $lines)
        Write-Info ('共通ヘッダ: ' + ($shown.ToArray() -join ', ') + '（OpenAPI 用設定。設定 common_headers の同じ名前のヘッダは置き換え）')
    }
    # アクセストークン: JMeter の実行中、API を呼び出すたびに TOKEN_DIR 直下の *.token のうち最新のファイルを読み込む
    $tokenDir = Get-OaConf $conf 'TOKEN_DIR'
    if ($tokenDir) {
        $raw['access_token_dir'] = $tokenDir
        $raw['access_token_file'] = $script:TokenPattern
        Write-Info ('アクセストークン: ' + $tokenDir + ' 直下の ' + $script:TokenPattern + ' のうち更新日時が最も新しいファイルを、JMeter の実行中に API の呼び出しごとに読み込みます')
        Check-TokenDir $tokenDir $script:TokenPattern $outDir
    }
}
function Check-TokenDir([string]$dirText, [string]$pattern, [string]$baseDir) {
    # 生成時の確認: JMeter の実行時と同じ規則でフォルダを決め（\ と / は同じ扱い、相対パスは JMX のフォルダ基準、
    # ; で区切った場合は最初に存在するフォルダ）、現在いちばん新しいファイルを表示する
    $isWin = [System.IO.Path]::DirectorySeparatorChar -eq [char]'\'
    $dirs = New-Object System.Collections.Generic.List[string]
    foreach ($d0 in $dirText.Split([char]';')) {
        $d = $d0.Trim().Replace('\', '/')
        if (-not $d) { continue }
        if ($d.StartsWith('~/', $script:Ord)) { $d = $d.Substring(2) }
        if ($isWin) { $abs = ($d -match '^[A-Za-z]:/') -or $d.StartsWith('//', $script:Ord) } else { $abs = $d.StartsWith('/', $script:Ord) }
        try {
            if (-not $abs) { $d = [System.IO.Path]::Combine((FullPath $baseDir), $d) }
            $d = [System.IO.Path]::GetFullPath($d)
        } catch { }
        $dirs.Add($d)
    }
    $found = $null
    foreach ($d in $dirs) { if ([System.IO.Directory]::Exists($d)) { $found = $d; break } }
    if ($null -eq $found) {
        Write-Warn ('アクセストークンのフォルダが見つかりません: ' + ($dirs.ToArray() -join ' / ') + '（JMeter を実行する環境に作成してください。相対パスは JMX のフォルダ基準）')
        return
    }
    try { $items = @(Get-ChildItem -LiteralPath $found -File -Force) }
    catch { Write-Warn ('アクセストークンのフォルダを読み取れません: ' + $found + '（' + $_.Exception.Message + '）'); return }
    $best = $null
    $bestMs = 0L
    $count = 0
    foreach ($f in $items) {
        if (-not ($f.Name -like $pattern)) { continue }
        $count++
        $ms = [long][Math]::Floor($f.LastWriteTimeUtc.Ticks / 10000)
        if ($null -eq $best -or $ms -gt $bestMs -or ($ms -eq $bestMs -and [string]::CompareOrdinal($f.Name, $best.Name) -gt 0)) { $best = $f; $bestMs = $ms }
    }
    if ($null -eq $best) {
        Write-Warn ('アクセストークンのフォルダに ' + $pattern + ' のファイルがありません: ' + $found + '（JMeter の実行までに配置してください）')
        return
    }
    Write-Info ('アクセストークン: 現在の最新ファイル ' + $best.FullName + '（更新日時 ' + $best.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss', $script:Inv) + '。' + $pattern + ' は ' + $count + ' 件）')
}
function Add-OaApiHeaders($records, $conf) {
    # 各 API のヘッダ（API ごとの HTTP Header Manager）の初期値: OpenAPI 用設定 API_HEADERS（-OpenApiConfig を指定しない場合・空欄は既定値、「なし」は付けない）。
    # OpenAPI 定義から求めた同じ名前のヘッダ（Accept など）は置き換え、API一覧の「ヘッダ」の後ろに追加する
    $text = ''
    if ($null -ne $conf) { $text = Get-OaConf $conf 'API_HEADERS' }
    $where = 'OpenAPI 用設定 API_HEADERS'
    if (-not $text) {
        foreach ($c in $script:Spec.openapi_config) { if ($c.key -ceq 'API_HEADERS') { $text = [string]$c.default } }
        $where += ' の既定値'
    }
    if ((Parse-OnOff $text) -ceq 'OFF') { return }
    $add = Parse-Headers $text $where
    if ($add.Count -eq 0) { return }
    $names = New-Object System.Collections.Generic.List[string]
    $shown = New-Object System.Collections.Generic.List[string]
    foreach ($h in $add) { $names.Add(([string]$h[0]).ToLowerInvariant()); $shown.Add([string]$h[0]) }
    foreach ($rec in $records) {
        $lines = New-Object System.Collections.Generic.List[string]
        foreach ($x in ([string]$rec['headers']).Split([char]10)) {
            if ($x.Trim() -and -not $names.Contains($x.Split([char]':', 2)[0].Trim().ToLowerInvariant())) { $lines.Add($x) }
        }
        foreach ($h in $add) { $lines.Add([string]$h[0] + ': ' + [string]$h[1]) }
        $rec['headers'] = [string]::Join("`n", $lines)
    }
    Write-Info ('各 API のヘッダ: ' + ($shown.ToArray() -join ', ') + '（' + $where + '。OpenAPI 定義から求めた同じ名前のヘッダは置き換え）')
}

# ----------------------------------------------------------------------------- input loading
function Get-Aliases {
    $m = New-Object 'System.Collections.Generic.Dictionary[string,string]'
    foreach ($c in $script:Spec.api_columns) {
        $all = New-Object System.Collections.Generic.List[string]
        $all.Add($c.key)
        $all.Add($c.header)
        foreach ($a in $c.aliases) { $all.Add($a) }
        foreach ($a in $all) {
            $k = NormalizeKey $a
            if (-not $m.ContainsKey($k)) { $m[$k] = $c.key }
        }
    }
    return ,$m
}
function Load-ApiTable($rows, [string]$source) {
    $alias = Get-Aliases
    $headerAt = -1
    $colmap = $null
    for ($idx = 0; $idx -lt $rows.Count; $idx++) {
        $cells = $rows[$idx][1]
        $found = New-Object 'System.Collections.Generic.Dictionary[string,int]'
        for ($ci = 0; $ci -lt $cells.Length; $ci++) {
            $nk = NormalizeKey $cells[$ci]
            if ($alias.ContainsKey($nk)) {
                $k = $alias[$nk]
                if (-not $found.ContainsKey($k)) { $found[$k] = $ci }
            }
        }
        if ($found.ContainsKey('url')) { $headerAt = $idx; $colmap = $found; break }
    }
    if ($headerAt -lt 0) { throw (New-ToolError ($source + ': 見出し行（「URL」列を含む行）が見つかりません')) }
    $records = New-Object System.Collections.Generic.List[object]
    for ($idx = $headerAt + 1; $idx -lt $rows.Count; $idx++) {
        $rn = $rows[$idx][0]
        $cells = $rows[$idx][1]
        $first = ''
        foreach ($v in $cells) { if ($v.Trim()) { $first = $v.Trim(); break } }
        if ($first -ceq '' -or $first.StartsWith('#', $script:Ord)) { continue }
        $rec = @{ _row = $rn }
        foreach ($c in $script:Spec.api_columns) {
            $ci = -1
            if ($colmap.ContainsKey($c.key)) { $ci = $colmap[$c.key] }
            if ($ci -ge 0 -and $ci -lt $cells.Length) { $rec[$c.key] = (Clean $cells[$ci]).Trim() } else { $rec[$c.key] = '' }
        }
        $records.Add($rec)
    }
    return ,$records
}
function Load-SettingsTable($rows, [string]$source) {
    $kc = -1
    $vc = -1
    $start = -1
    for ($idx = 0; $idx -lt $rows.Count; $idx++) {
        $cells = $rows[$idx][1]
        $kc = -1
        $vc = -1
        for ($ci = 0; $ci -lt $cells.Length; $ci++) {
            $nv = NormalizeKey $cells[$ci]
            if ($kc -lt 0 -and ($nv -ceq '設定キー' -or $nv -ceq 'key' -or $nv -ceq 'キー')) { $kc = $ci }
            if ($vc -lt 0 -and ($nv -ceq '設定値' -or $nv -ceq 'value' -or $nv -ceq '値')) { $vc = $ci }
        }
        if ($kc -ge 0 -and $vc -ge 0) { $start = $idx; break }
    }
    if ($start -lt 0) { throw (New-ToolError ($source + ': 見出し行（「設定キー」「設定値」）が見つかりません')) }
    $values = New-Object 'System.Collections.Generic.Dictionary[string,string]'
    for ($idx = $start + 1; $idx -lt $rows.Count; $idx++) {
        $cells = $rows[$idx][1]
        $key = ''
        if ($kc -lt $cells.Length) { $key = $cells[$kc].Trim() }
        if (-not [regex]::IsMatch($key, '^[A-Za-z_][A-Za-z0-9_]*$')) { continue }
        $val = ''
        if ($vc -lt $cells.Length) { $val = (Clean $cells[$vc]).Trim() }
        $values[$key] = $val
    }
    return ,$values
}
function Find-Sheet($sheets, [string[]]$names) {
    $want = New-Object System.Collections.Generic.List[string]
    foreach ($x in $names) { $want.Add((NormalizeKey $x)) }
    foreach ($s in $sheets) { if ($want.Contains((NormalizeKey $s[0]))) { return ,$s } }
    return $null
}
function Rows-FromCsv($table) {
    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $table.Count; $i++) { $rows.Add(@(($i + 1), [string[]]$table[$i])) }
    return ,$rows
}
function Load-Input([string]$inputPath, [string]$settingsPath) {
    $ext = [System.IO.Path]::GetExtension($inputPath).ToLowerInvariant()
    $raw = $null
    $oaDefaults = @()
    $oaConf = $null
    if ($OpenApiConfig -and $ext -cne '.json') {
        Write-Warn ('OpenAPI 用設定ファイル（-OpenApiConfig）は OpenAPI 定義（.json）を入力する場合だけ使います（無視します）: ' + $OpenApiConfig)
    }
    $base = [System.IO.Path]::GetFileName($inputPath)
    if ($ext -ceq '.xlsx' -or $ext -ceq '.xlsm') {
        $sheets = Read-Xlsx $inputPath
        $found = Find-Sheet $sheets @($script:Tpl.api_sheet, 'API', 'APIs', 'api_list')
        if ($null -eq $found) {
            foreach ($s in $sheets) {
                try { [void](Load-ApiTable $s[1] $s[0]); $found = $s; break } catch [System.InvalidOperationException] { continue }
            }
        }
        if ($null -eq $found) { throw (New-ToolError ('「' + $script:Tpl.api_sheet + '」シートが見つかりません: ' + $inputPath)) }
        $records = Load-ApiTable $found[1] ($base + '（' + $found[0] + 'シート）')
        if (-not $settingsPath) {
            $sf = Find-Sheet $sheets @($script:Tpl.settings_sheet, 'settings')
            if ($null -ne $sf) {
                $raw = Load-SettingsTable $sf[1] ($base + '（' + $sf[0] + 'シート）')
                Write-Info ('設定: 「' + $sf[0] + '」シートを使用します')
            } else {
                Write-Warn ('「' + $script:Tpl.settings_sheet + '」シートが無いため、設定はすべて既定値を使用します')
            }
        }
        Write-Info ('API定義: 「' + $found[0] + '」シート（' + $records.Count + ' 件）')
    } elseif ($ext -ceq '.csv' -or $ext -ceq '.txt') {
        $tx = Read-TextAuto $inputPath
        $records = Load-ApiTable (Rows-FromCsv (Parse-Csv $tx[0])) $base
        Write-Info ('API定義: CSV（文字コード ' + $tx[1] + '、' + $records.Count + ' 件）')
        if (-not $settingsPath) {
            $cand = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($inputPath), [System.IO.Path]::GetFileNameWithoutExtension($inputPath) + '_settings.csv')
            if (Test-Path -LiteralPath $cand -PathType Leaf) { $settingsPath = $cand }
            else { Write-Warn ('設定 CSV（' + [System.IO.Path]::GetFileName($cand) + '）が無いため、設定はすべて既定値を使用します') }
        }
    } elseif ($ext -ceq '.json') {
        if ($OpenApiConfig) {
            $oaConf = Load-OaConfig $OpenApiConfig
            $script:OaAuth = [bool](Get-OaConf $oaConf 'TOKEN_DIR') -or (((Get-OaConf $oaConf 'HEADERS') + "`n" + (Get-OaConf $oaConf 'API_HEADERS')) -match '(?im)^\s*authorization\s*:')
            $script:OaConf = $oaConf
        }
        $oa = Load-OpenApi $inputPath
        $records = $oa[0]
        $oaDefaults = $oa[1]
        Add-OaApiHeaders $records $oaConf
        if (-not $settingsPath) {
            $cand = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($inputPath), [System.IO.Path]::GetFileNameWithoutExtension($inputPath) + '_settings.csv')
            if (Test-Path -LiteralPath $cand -PathType Leaf) { $settingsPath = $cand }
            else { Write-Info ('設定 CSV（' + [System.IO.Path]::GetFileName($cand) + '）が無いため既定値を使用します（接続先などは OpenAPI 定義の値）') }
        }
    } elseif ($ext -ceq '.xls') {
        throw (New-ToolError ('.xls 形式（Excel 97-2003）には対応していません。Excel で .xlsx 形式で保存し直してください: ' + $inputPath))
    } else {
        throw (New-ToolError ('対応していないファイル形式です（.xlsx / .csv / .json）: ' + $inputPath))
    }
    if ($settingsPath) {
        if (-not (Test-Path -LiteralPath $settingsPath -PathType Leaf)) { throw (New-ToolError ('設定ファイルが見つかりません: ' + $settingsPath)) }
        $sext = [System.IO.Path]::GetExtension($settingsPath).ToLowerInvariant()
        if ($sext -ceq '.xlsx' -or $sext -ceq '.xlsm') {
            $sheets = Read-Xlsx $settingsPath
            $sf = Find-Sheet $sheets @($script:Tpl.settings_sheet, 'settings')
            if ($null -eq $sf) { throw (New-ToolError ('「' + $script:Tpl.settings_sheet + '」シートが見つかりません: ' + $settingsPath)) }
            $raw = Load-SettingsTable $sf[1] ([System.IO.Path]::GetFileName($settingsPath))
        } else {
            $tx = Read-TextAuto $settingsPath
            $raw = Load-SettingsTable (Rows-FromCsv (Parse-Csv $tx[0])) ([System.IO.Path]::GetFileName($settingsPath))
        }
        Write-Info ('設定: ' + $settingsPath)
    }
    if ($null -eq $raw) { $raw = New-Object 'System.Collections.Generic.Dictionary[string,string]' }
    foreach ($d in $oaDefaults) {
        if (-not $raw.ContainsKey($d[0]) -or -not $raw[$d[0]]) {
            $raw[$d[0]] = [string]$d[1]
            Write-Info ('設定 ' + $d[0] + ': ' + $d[1] + '（設定ファイルで未指定のため OpenAPI 定義から設定）')
        }
    }
    if ($null -ne $oaConf) {
        $outBase = $OutDir
        if (-not $outBase) { $outBase = 'output' }
        Apply-OaConfig $raw $oaConf $outBase
    }
    return @($raw, $records)
}
function Parse-Bool([string]$v) {
    $k = NormalizeKey $v
    if (@('true', 'yes', 'y', '1', 'on', 'はい', '有効', 'する') -ccontains $k) { return $true }
    if (@('false', 'no', 'n', '0', 'off', 'いいえ', '無効', 'しない') -ccontains $k) { return $false }
    return $null
}
function Parse-OnOff([string]$v) {
    $k = NormalizeKey $v
    if (@('on', 'y', 'yes', 'true', '1', '有効', '生成') -ccontains $k) { return 'ON' }
    if (@('off', 'n', 'no', 'false', '0', 'なし', '生成しない', '不要') -ccontains $k) { return 'OFF' }
    if (@('disable', 'disabled', '無効', '無効化') -ccontains $k) { return 'DISABLE' }
    return $null
}
function Resolve-Settings($raw) {
    $known = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($s in $script:Spec.settings) { [void]$known.Add($s.key) }
    $rk = New-Object System.Collections.Generic.List[string]
    foreach ($k in $raw.Keys) { $rk.Add($k) }
    $rk.Sort([System.StringComparer]::Ordinal)
    foreach ($k in $rk) { if (-not $known.Contains($k)) { Write-Warn ('設定: 不明な設定キー「' + $k + '」は無視します') } }
    $cfg = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    foreach ($s in $script:Spec.settings) {
        $k = [string]$s.key
        $t = [string]$s.type
        $def = [string]$s.default
        $v = ''
        if ($raw.ContainsKey($k)) { $v = [string]$raw[$k] }
        if ($v -ceq '') { $v = $def }
        if ($t -ceq 'int') {
            if ($v -cne '' -and -not [regex]::IsMatch($v, '^-?[0-9]+$')) {
                Write-Warn ('設定 ' + $k + ': 「' + $v + '」は数値ではないため既定値「' + $def + '」を使用します')
                $v = $def
            }
            $cfg[$k] = $v
        } elseif ($t -ceq 'bool') {
            $b = Parse-Bool $v
            if ($null -eq $b) {
                Write-Warn ('設定 ' + $k + ': 「' + $v + '」は true / false ではないため既定値「' + $def + '」を使用します')
                $b = Parse-Bool $def
            }
            $cfg[$k] = [bool]$b
        } elseif ($t -ceq 'onoff') {
            $o = Parse-OnOff $v
            if ($null -eq $o) {
                Write-Warn ('設定 ' + $k + ': 「' + $v + '」は ON / OFF / DISABLE ではないため既定値「' + $def + '」を使用します')
                $o = Parse-OnOff $def
            }
            $cfg[$k] = $o
        } elseif ($t -ceq 'enum') {
            $lv = $v.ToLowerInvariant()
            if (-not (@($s.choices) -ccontains $lv)) {
                Write-Warn ('設定 ' + $k + ': 「' + $v + '」は ' + (@($s.choices) -join ' / ') + ' のいずれでもないため既定値「' + $def + '」を使用します')
                $lv = $def
            }
            $cfg[$k] = $lv
        } else {
            $cfg[$k] = $v
        }
    }
    $cfg['_common_headers'] = Parse-Headers $cfg['common_headers'] '設定 common_headers'
    if ([string]$cfg['access_token_dir'] -cne '') {
        $hasAuth = $false
        foreach ($h in $cfg['_common_headers']) { if (([string]$h[0]).ToLowerInvariant() -ceq 'authorization') { $hasAuth = $true } }
        if (-not $hasAuth) { $cfg['_common_headers'].Add(@('Authorization', 'Bearer ${ACCESS_TOKEN}')) }
    }
    $cfg['_user_vars'] = Parse-PairsStrict $cfg['user_variables'] '設定 user_variables'
    $cfg['_cookies'] = Parse-PairsStrict $cfg['cookies'] '設定 cookies'
    if ([string]$cfg['duration'] -cne '' -and [long]$cfg['duration'] -gt 0 -and [string]$cfg['loops'] -cne '-1') {
        Write-Warn ('設定: 実行時間（duration）を指定する場合は、ループ回数（loops）を -1 にしてください（現在: ' + $cfg['loops'] + '）')
    }
    return ,$cfg
}

# ----------------------------------------------------------------------------- API normalisation
$script:UrlRe = '^([A-Za-z][A-Za-z0-9+.\-]*)://([^/?#]*)([^?#]*)(?:\?([^#]*))?(?:#.*)?$'
$script:VarRe = '^[A-Za-z_][A-Za-z0-9_]*$'
$script:Unreserved = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~'
$script:PathSafe = $script:Unreserved + "/:@!`$&'()*+,;="
$script:QueryInArgs = @('GET', 'DELETE', 'OPTIONS')
$script:NoBody = @('GET', 'DELETE', 'OPTIONS', 'HEAD')

function Parse-Url([string]$url) {
    $m = [regex]::Match($url, $script:UrlRe)
    if (-not $m.Success) { return $null }
    $auth = $m.Groups[2].Value
    $hostName = $auth
    $port = ''
    $pm = [regex]::Match($auth, '^(.*):([0-9]+|\$\{[^}]*\})$')
    if ($pm.Success) { $hostName = $pm.Groups[1].Value; $port = $pm.Groups[2].Value }
    if ($hostName.Contains('@')) { $parts = $hostName.Split('@'); $hostName = $parts[$parts.Length - 1] }
    $path = $m.Groups[3].Value
    if ($path -ceq '') { $path = '/' }
    $query = ''
    if ($m.Groups[4].Success) { $query = $m.Groups[4].Value }
    return @{ scheme = $m.Groups[1].Value.ToLowerInvariant(); host = $hostName; port = $port; path = $path; query = $query }
}
function Close-Brace([string]$s, [int]$i) {
    $depth = 0
    $j = $i + 1
    while ($j -lt $s.Length) {
        if ($s[$j] -ceq '{') { $depth++ }
        elseif ($s[$j] -ceq '}') { $depth--; if ($depth -eq 0) { return $j } }
        $j++
    }
    return -1
}
function Pct-Encode([string]$s, [string]$safe) {
    $out = New-Object System.Text.StringBuilder
    $i = 0
    $n = $s.Length
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    while ($i -lt $n) {
        $ch = $s[$i]
        if ($ch -ceq '$' -and $i + 1 -lt $n -and $s[$i + 1] -ceq '{') {
            $j = Close-Brace $s $i
            if ($j -gt 0) { [void]$out.Append($s.Substring($i, $j + 1 - $i)); $i = $j + 1; continue }
        }
        if ($ch -ceq '%' -and $i + 2 -lt $n + 0 -and [regex]::IsMatch($s.Substring($i, [Math]::Min(3, $n - $i)), '^%[0-9A-Fa-f]{2}')) {
            [void]$out.Append($s.Substring($i, 3)); $i += 3; continue
        }
        if ($safe.IndexOf($ch) -ge 0) { [void]$out.Append($ch); $i++; continue }
        $len = 1
        if ([char]::IsHighSurrogate($ch) -and $i + 1 -lt $n -and [char]::IsLowSurrogate($s[$i + 1])) { $len = 2 }
        foreach ($b in $utf8.GetBytes($s.Substring($i, $len))) { [void]$out.Append('%' + ([int]$b).ToString('X2', $script:Inv)) }
        $i += $len
    }
    return $out.ToString()
}
function Needs-Encode([string]$v) { return -not [regex]::IsMatch($v, '%[0-9A-Fa-f]{2}') }
function Split-Pairs([string]$text) {
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($line in (NormNl $text).Split([char]10)) {
        foreach ($part in $line.Split([char]'&')) {
            $p = $part.Trim()
            if (-not $p) { continue }
            $eq = $p.IndexOf('=')
            if ($eq -ge 0) { $out.Add(@($p.Substring(0, $eq).Trim(), $p.Substring($eq + 1).Trim())) } else { $out.Add(@($p, '')) }
        }
    }
    return ,$out
}
function Parse-Headers([string]$text, [string]$where) {
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($line in (NormNl $text).Split([char]10)) {
        $s = $line.Trim()
        if (-not $s) { continue }
        $c = $s.IndexOf(':')
        if ($c -lt 0 -or -not $s.Substring(0, $c).Trim()) {
            Write-Warn ($where + ': ヘッダは「名前: 値」の形式で記入してください（無視します）: ' + $s)
            continue
        }
        $out.Add(@($s.Substring(0, $c).Trim(), $s.Substring($c + 1).Trim()))
    }
    return ,$out
}
function Parse-PairsStrict([string]$text, [string]$where) {
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($line in (NormNl $text).Split([char]10)) {
        $s = $line.Trim()
        if (-not $s) { continue }
        $eq = $s.IndexOf('=')
        if ($eq -lt 0) { Write-Warn ($where + ': 「名前=値」の形式で記入してください（無視します）: ' + $s); continue }
        $k = $s.Substring(0, $eq).Trim()
        if (-not $k) { Write-Warn ($where + ': 名前が空です（無視します）: ' + $s); continue }
        $out.Add(@($k, $s.Substring($eq + 1).Trim()))
    }
    return ,$out
}
function Parse-Assign([string]$text, [string]$where, [string]$what) {
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($line in (NormNl $text).Split([char]10)) {
        $s = $line.Trim()
        if (-not $s) { continue }
        $eq = $s.IndexOf('=')
        $k = ''
        $v = ''
        if ($eq -ge 0) { $k = $s.Substring(0, $eq).Trim(); $v = $s.Substring($eq + 1).Trim() }
        if ($eq -lt 0 -or -not [regex]::IsMatch($k, $script:VarRe) -or -not $v) {
            Write-Warn ($where + ': ' + $what + 'は「変数名=式」の形式で記入してください（無視します）: ' + $s)
            continue
        }
        $out.Add(@($k, $v))
    }
    return ,$out
}
function Is-Form([string]$body) {
    $segs = New-Object System.Collections.Generic.List[string]
    foreach ($x in [regex]::Split($body, '[\n&]')) { if ($x.Trim()) { $segs.Add($x.Trim()) } }
    if ($segs.Count -eq 0) { return $false }
    foreach ($x in $segs) { if (-not [regex]::IsMatch($x, '^[^=\s]+=')) { return $false } }
    return $true
}
function Parse-Num([string]$v, [string]$label, [string]$where) {
    if ($v -ceq '') { return -1 }
    if ([regex]::IsMatch($v, '^[0-9]+$')) { return [long]$v }
    Write-Warn ($where + ': ' + $label + '「' + $v + '」は数値で記入してください（無視します）')
    return -1
}
function Normalize-Api($rec, [int]$index, $cfg) {
    $where = [string]$rec._row + '行目'
    $no = [string]$rec['no']
    if ([regex]::IsMatch($no, '^[0-9]{1,6}$')) { $aid = 'API' + ([int]$no).ToString('00', $script:Inv) }
    elseif ($no) { $aid = 'API_' + [regex]::Replace($no, '[^A-Za-z0-9]', '_') }
    else { $aid = 'API' + ($index + 1).ToString('00', $script:Inv) }
    $where = $where + '（' + $aid + '）'
    $url = [string]$rec['url']
    if (-not $url) { Write-Warn ($where + ': URL が空のためスキップします'); return $null }
    $enabled = $true
    if ($rec['enabled']) {
        $e = Parse-Bool $rec['enabled']
        if ($null -eq $e) { Write-Warn ($where + ': 有効「' + $rec['enabled'] + '」は Y / N で記入してください（Y として扱います）') }
        else { $enabled = [bool]$e }
    }
    $method = ([string]$rec['method']).ToUpperInvariant()
    if (-not $method) { Write-Warn ($where + ': メソッドが空のため GET として扱います'); $method = 'GET' }
    if (-not (@('GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'HEAD', 'OPTIONS') -ccontains $method)) {
        Write-Warn ($where + ': メソッド「' + $method + '」はそのまま設定します（GET / POST 以外は動作を確認してください）')
    }
    $full = $url
    if ($url.StartsWith('/', $script:Ord)) {
        if (-not $cfg['base_url']) { Write-Warn ($where + ': URL が「/」で始まっていますが設定 base_url が空のためスキップします'); return $null }
        $full = ([string]$cfg['base_url']).TrimEnd('/') + $url
    }
    $p = Parse-Url $full
    if ($null -eq $p) {
        $p = Parse-Url ('https://' + $full)
        if ($null -eq $p) { Write-Warn ($where + ': URL の形式が正しくないためスキップします: ' + $url); return $null }
        Write-Warn ($where + ': URL にスキーム（https://）が無いため https として扱います')
        $full = 'https://' + $full
    }
    $pairs = New-Object System.Collections.Generic.List[object]
    foreach ($x in (Split-Pairs $p.query)) { $pairs.Add($x) }
    foreach ($x in (Split-Pairs $rec['query'])) { $pairs.Add($x) }
    $path = Pct-Encode $p.path $script:PathSafe
    $args2 = New-Object System.Collections.Generic.List[object]
    if ($script:QueryInArgs -ccontains $method) {
        foreach ($x in $pairs) { $args2.Add(@($x[0], $x[1], (Needs-Encode $x[1]))) }
    } elseif ($pairs.Count -gt 0) {
        $qs = New-Object System.Collections.Generic.List[string]
        foreach ($x in $pairs) { $qs.Add((Pct-Encode $x[0] $script:Unreserved) + '=' + (Pct-Encode $x[1] $script:Unreserved)) }
        $path += '?' + ($qs.ToArray() -join '&')
    }
    $headers = Parse-Headers $rec['headers'] $where
    $body = [string]$rec['body']
    $bt = ([string]$rec['body_type']).ToLowerInvariant()
    if (-not $bt) { $bt = 'auto' }
    if (-not (@('auto', 'json', 'form', 'raw', 'xml', 'multipart') -ccontains $bt)) {
        Write-Warn ($where + ': ボディ形式「' + $rec['body_type'] + '」は auto / json / form / raw のいずれかで記入してください（auto として扱います）')
        $bt = 'auto'
    }
    $kind = ''
    if ($body -and ($script:NoBody -ccontains $method)) {
        Write-Warn ($where + ': ' + $method + ' ではリクエストボディを使用できないため無視します')
        $body = ''
    }
    if ($body) {
        if ($bt -ceq 'auto') {
            $s = $body.TrimStart()
            if ($s.StartsWith('{', $script:Ord) -or $s.StartsWith('[', $script:Ord)) { $kind = 'json' }
            elseif ($s.StartsWith('<', $script:Ord)) { $kind = 'xml' }
            elseif (Is-Form $body) { $kind = 'form' }
            else { $kind = 'raw' }
        } else {
            $kind = $bt
        }
    }
    if ($kind -ceq 'form' -or $kind -ceq 'multipart') {
        $args2 = New-Object System.Collections.Generic.List[object]
        foreach ($x in (Split-Pairs $body)) { $args2.Add(@($x[0], $x[1], (Needs-Encode $x[1]))) }
    }
    if ($kind -ceq 'json') { $body = [regex]::Replace($body, '("(?:[^"\\]|\\.)*")|\u3000', { param($m) if ($m.Groups[1].Success) { $m.Value } else { ' ' } }) }
    $ctype = ''
    if ($kind -ceq 'json') { $ctype = 'application/json' } elseif ($kind -ceq 'xml') { $ctype = 'application/xml' }
    if ($ctype) {
        $hasCt = $false
        foreach ($h in $headers) { if (([string]$h[0]).ToLowerInvariant() -ceq 'content-type') { $hasCt = $true } }
        foreach ($h in $cfg['_common_headers']) { if (([string]$h[0]).ToLowerInvariant() -ceq 'content-type') { $hasCt = $true } }
        if (-not $hasCt) { $headers.Add(@('Content-Type', $ctype)) }
    }
    $wait = Parse-Num $rec['wait_ms'] '待機時間' $where
    if ($wait -lt 0) {
        if ([regex]::IsMatch([string]$cfg['default_wait_ms'], '^[0-9]+$')) { $wait = [long]$cfg['default_wait_ms'] } else { $wait = 0 }
    }
    $maxMs = Parse-Num $rec['max_time_ms'] '応答時間上限' $where
    $statuses = New-Object System.Collections.Generic.List[string]
    foreach ($c in [regex]::Split([string]$rec['expected_status'], '[,\s/]+')) {
        if (-not $c) { continue }
        if ([regex]::IsMatch($c, '^[1-5][0-9][0-9]$')) { if (-not $statuses.Contains($c)) { $statuses.Add($c) } }
        else { Write-Warn ($where + ': 期待ステータス「' + $c + '」は 3 桁の数字で記入してください（無視します）') }
    }
    $texts = New-Object System.Collections.Generic.List[string]
    foreach ($x in (NormNl $rec['expected_text']).Split([char]10)) { if ($x.Trim()) { $texts.Add($x) } }
    $name = ([regex]::Replace([string]$rec['name'], '\s+', ' ')).Trim()
    if (-not $name) { $name = $method + ' ' + $p.path }
    $rawBody = ''
    if ($kind -ceq 'json' -or $kind -ceq 'xml' -or $kind -ceq 'raw') { $rawBody = $body }
    return @{
        row = $rec._row; id = $aid; label = ($aid + '_' + $name); enabled = $enabled; method = $method
        url = $full; scheme = $p.scheme; host = $p.host; port = $p.port; path = $path
        args = $args2; body = $rawBody; kind = $kind; headers = $headers
        wait = $wait; max_ms = $maxMs; statuses = $statuses; texts = $texts
        regex = (Parse-Assign $rec['regex_extract'] $where '正規表現抽出')
        json = (Parse-Assign $rec['json_extract'] $where 'JSON抽出')
        pre = [string]$rec['pre_script']; post = [string]$rec['post_script']; comment = [string]$rec['comment']
        hostref = $null
    }
}

# ----------------------------------------------------------------------------- JMX builder
function XEl([string]$tag, $attrs, $children, $text) {
    $a = New-Object System.Collections.Generic.List[object]
    if ($null -ne $attrs) { foreach ($x in $attrs) { $a.Add($x) } }
    $k = New-Object System.Collections.Generic.List[object]
    if ($null -ne $children) { foreach ($x in $children) { $k.Add($x) } }
    return @{ tag = $tag; attrs = $a; kids = $k; text = $text }
}
function PStr([string]$name, [string]$value) { return (XEl 'stringProp' @(, @('name', $name)) $null $value) }
function PBool([string]$name, [bool]$value) { $t = 'false'; if ($value) { $t = 'true' }; return (XEl 'boolProp' @(, @('name', $name)) $null $t) }
function PInt([string]$name, $value) { return (XEl 'intProp' @(, @('name', $name)) $null ([string]$value)) }
function PLong([string]$name, $value) { return (XEl 'longProp' @(, @('name', $name)) $null ([string]$value)) }
function PColl([string]$name, $items) { return (XEl 'collectionProp' @(, @('name', $name)) $items $null) }
function PElem([string]$name, [string]$etype, $children, $extra) {
    $attrs = New-Object System.Collections.Generic.List[object]
    $attrs.Add(@('name', $name))
    $attrs.Add(@('elementType', $etype))
    if ($null -ne $extra) { foreach ($x in $extra) { $attrs.Add($x) } }
    return (XEl 'elementProp' $attrs $children $null)
}
function JObj([string]$tag, [string]$gui, [string]$cls, [string]$name, [bool]$enabled, $props, [string]$comment) {
    $kids = New-Object System.Collections.Generic.List[object]
    if ($comment) { $kids.Add((PStr 'TestPlan.comments' $comment)) }
    foreach ($p in $props) { $kids.Add($p) }
    $en = 'false'
    if ($enabled) { $en = 'true' }
    return (XEl $tag @(@('guiclass', $gui), @('testclass', $cls), @('testname', $name), @('enabled', $en)) $kids $null)
}
function JNode($elem, $children) {
    $k = New-Object System.Collections.Generic.List[object]
    if ($null -ne $children) { foreach ($x in $children) { $k.Add($x) } }
    return @{ elem = $elem; kids = $k }
}
function PEsc([string]$v) { return $v.Replace('\', '\\').Replace(',', '\,') }
function PFun([string]$name, [string]$default) { return '${__P(' + $name + ',' + (PEsc $default) + ')}' }
function Ser-Elem($e, [int]$ind, $out) {
    $pad = '  ' * $ind
    $a = New-Object System.Text.StringBuilder
    foreach ($kv in $e.attrs) { [void]$a.Append(' ' + $kv[0] + '="' + (XEsc ([string]$kv[1])) + '"') }
    if ($e.kids.Count -gt 0) {
        [void]$out.Append($pad + '<' + $e.tag + $a.ToString() + ">`n")
        foreach ($c in $e.kids) { Ser-Elem $c ($ind + 1) $out }
        [void]$out.Append($pad + '</' + $e.tag + ">`n")
    } elseif ($null -eq $e.text) {
        [void]$out.Append($pad + '<' + $e.tag + $a.ToString() + "/>`n")
    } else {
        [void]$out.Append($pad + '<' + $e.tag + $a.ToString() + '>' + (XEsc ([string]$e.text)) + '</' + $e.tag + ">`n")
    }
}
function Ser-Nodes($nodes, [int]$ind, $out) {
    foreach ($nd in $nodes) {
        Ser-Elem $nd.elem $ind $out
        if ($nd.kids.Count -gt 0) {
            [void]$out.Append(('  ' * $ind) + "<hashTree>`n")
            Ser-Nodes $nd.kids ($ind + 1) $out
            [void]$out.Append(('  ' * $ind) + "</hashTree>`n")
        } else {
            [void]$out.Append(('  ' * $ind) + "<hashTree/>`n")
        }
    }
}
function Jmx-Text($plan) {
    $out = New-Object System.Text.StringBuilder
    [void]$out.Append("<?xml version=`"1.0`" encoding=`"UTF-8`"?>`n")
    [void]$out.Append('<jmeterTestPlan version="1.2" properties="5.0" jmeter="' + $script:Tool.jmeter + "`">`n")
    [void]$out.Append("  <hashTree>`n")
    Ser-Nodes @(, $plan) 2 $out
    [void]$out.Append("  </hashTree>`n</jmeterTestPlan>`n")
    return $out.ToString()
}
function Jsr223([string]$tag, [string]$name, [bool]$enabled, [string]$params, [string]$script, [string]$comment) {
    return (JObj $tag 'TestBeanGUI' $tag $name $enabled @((PStr 'scriptLanguage' 'groovy'), (PStr 'parameters' $params), (PStr 'filename' ''),
            (PStr 'cacheKey' 'true'), (PStr 'script' $script)) $comment)
}
function Header-Manager([string]$name, $headers, [string]$comment) {
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($h in $headers) { $items.Add((PElem '' 'Header' @((PStr 'Header.name' $h[0]), (PStr 'Header.value' $h[1])) $null)) }
    return (JObj 'HeaderManager' 'HeaderPanel' 'HeaderManager' $name $true @(, (PColl 'HeaderManager.headers' $items)) $comment)
}
function Thread-Group([string]$tag, [string]$gui, [string]$name, [bool]$enabled, [string]$threads, [string]$rampup, [string]$loops, [bool]$scheduler, [string]$duration, [string]$onError, [string]$comment) {
    $loop = PElem 'ThreadGroup.main_controller' 'LoopController' @((PBool 'LoopController.continue_forever' $false), (PStr 'LoopController.loops' $loops)) @(@('guiclass', 'LoopControlPanel'), @('testclass', 'LoopController'), @('testname', 'Loop Controller'), @('enabled', 'true'))
    return (JObj $tag $gui $tag $name $enabled @((PStr 'ThreadGroup.on_sample_error' $onError), $loop, (PStr 'ThreadGroup.num_threads' $threads),
            (PStr 'ThreadGroup.ramp_time' $rampup), (PBool 'ThreadGroup.scheduler' $scheduler), (PStr 'ThreadGroup.duration' $duration),
            (PStr 'ThreadGroup.delay' ''), (PBool 'ThreadGroup.same_user_on_next_iteration' $true)) $comment)
}
function Save-Config([bool]$detail) {
    $fields = @(@('time', $true), @('latency', $true), @('timestamp', $true), @('success', $true), @('label', $true),
        @('code', $true), @('message', $true), @('threadName', $true), @('dataType', $true), @('encoding', $detail),
        @('assertions', $true), @('subresults', $detail), @('responseData', $detail), @('samplerData', $detail),
        @('xml', $detail), @('fieldNames', $true), @('responseHeaders', $detail), @('requestHeaders', $detail),
        @('responseDataOnError', $false), @('saveAssertionResultsFailureMessage', $true),
        @('assertionsResultsToSave', '0'), @('bytes', $true), @('sentBytes', $true), @('url', $true),
        @('threadCounts', $true), @('idleTime', $true), @('connectTime', $true))
    $kids = New-Object System.Collections.Generic.List[object]
    foreach ($f in $fields) {
        $v = $f[1]
        if ($v -is [string]) { $t = $v } elseif ($v) { $t = 'true' } else { $t = 'false' }
        $kids.Add((XEl $f[0] $null $null $t))
    }
    return (XEl 'objProp' $null @((XEl 'name' $null $null 'saveConfig'), (XEl 'value' @(, @('class', 'SampleSaveConfiguration')) $kids $null)) $null)
}
function Collector([string]$gui, [string]$name, [bool]$enabled, [string]$filename, [string]$comment, [bool]$errorsOnly = $false, [bool]$detail = $false) {
    return (JObj 'ResultCollector' $gui 'ResultCollector' $name $enabled @((PBool 'ResultCollector.error_logging' $errorsOnly), (Save-Config $detail), (PStr 'filename' $filename)) $comment)
}
function Assign-Hosts($cfg, $apis) {
    $hosts = New-Object System.Collections.Generic.List[object]
    $add = {
        param($scheme, $hostName, $port)
        foreach ($h in $hosts) { if ($h.scheme -ceq $scheme -and $h.host -ceq $hostName -and $h.port -ceq $port) { return $h } }
        $sfx = ''
        if ($hosts.Count -gt 0) { $sfx = '_' + ($hosts.Count + 1) }
        $h = @{ scheme = $scheme; host = $hostName; port = $port; pv = 'PROTOCOL' + $sfx; hv = 'HOST' + $sfx; portv = 'PORT' + $sfx }
        $hosts.Add($h)
        return $h
    }
    if ($cfg['base_url']) {
        $p = Parse-Url $cfg['base_url']
        if ($null -ne $p -and -not ($p.scheme + $p.host + $p.port).Contains('${')) { [void](& $add $p.scheme $p.host $p.port) }
    }
    foreach ($a in $apis) {
        if (($a.scheme + $a.host + $a.port).Contains('${')) { $a.hostref = $null }
        else { $a.hostref = (& $add $a.scheme $a.host $a.port) }
    }
    return ,$hosts
}
function Api-Node($api, $cfg) {
    $N = $script:Spec.names
    $CM = $script:Spec.comments
    $aid = $api.id
    $h = $api.hostref
    $props = New-Object System.Collections.Generic.List[object]
    if ($api.body) {
        $props.Add((PBool 'HTTPSampler.postBodyRaw' $true))
        $props.Add((PElem 'HTTPsampler.Arguments' 'Arguments' @(, (PColl 'Arguments.arguments' @(, (PElem '' 'HTTPArgument' @((PBool 'HTTPArgument.always_encode' $false), (PStr 'Argument.value' $api.body), (PStr 'Argument.metadata' '=')) $null)))) $null))
    } else {
        $items = New-Object System.Collections.Generic.List[object]
        foreach ($x in $api.args) {
            $items.Add((PElem $x[0] 'HTTPArgument' @((PBool 'HTTPArgument.always_encode' $x[2]), (PStr 'Argument.value' $x[1]), (PStr 'Argument.metadata' '='),
                        (PBool 'HTTPArgument.use_equals' $true), (PStr 'Argument.name' $x[0])) $null))
        }
        $props.Add((PElem 'HTTPsampler.Arguments' 'Arguments' @(, (PColl 'Arguments.arguments' $items)) @(@('guiclass', 'HTTPArgumentsPanel'), @('testclass', 'Arguments'), @('testname', 'User Defined Variables'), @('enabled', 'true'))))
    }
    if ($null -ne $h) { $dom = '${' + $h.hv + '}'; $prt = '${' + $h.portv + '}'; $pro = '${' + $h.pv + '}' }
    else { $dom = $api.host; $prt = $api.port; $pro = $api.scheme }
    foreach ($x in @((PStr 'HTTPSampler.domain' $dom), (PStr 'HTTPSampler.port' $prt), (PStr 'HTTPSampler.protocol' $pro),
            (PStr 'HTTPSampler.contentEncoding' $cfg['content_encoding']), (PStr 'HTTPSampler.path' $api.path), (PStr 'HTTPSampler.method' $api.method),
            (PBool 'HTTPSampler.follow_redirects' $cfg['follow_redirects']), (PBool 'HTTPSampler.auto_redirects' $false),
            (PBool 'HTTPSampler.use_keepalive' $cfg['keep_alive']), (PBool 'HTTPSampler.DO_MULTIPART_POST' ($api.kind -ceq 'multipart')),
            (PStr 'HTTPSampler.embedded_url_re' ''), (PStr 'HTTPSampler.connect_timeout' ''), (PStr 'HTTPSampler.response_timeout' ''))) { $props.Add($x) }
    $comment = Fill $CM.sampler @{ ID = $aid; METHOD = $api.method; URL = $api.url }
    if ($api.comment) { $comment += "`n" + $api.comment }
    $kids = New-Object System.Collections.Generic.List[object]
    $ids = @{ ID = $aid }
    if ($api.headers.Count -gt 0) { $kids.Add((JNode (Header-Manager (Fill $N.api_headers $ids) $api.headers $CM.api_headers) $null)) }
    if ($api.wait -gt 0) {
        $ms = [string]$api.wait
        $kids.Add((JNode (JObj 'ConstantTimer' 'ConstantTimerGui' 'ConstantTimer' (Fill $N.api_timer @{ ID = $aid; MS = $ms }) $true @(, (PStr 'ConstantTimer.delay' $ms)) (Fill $CM.api_timer @{ MS = $ms })) $null))
    }
    if ($api.pre) { $kids.Add((JNode (Jsr223 'JSR223PreProcessor' (Fill $N.api_pre $ids) $true '' $api.pre $CM.api_pre) $null)) }
    foreach ($x in $api.regex) {
        $var = $x[0]
        $expr = $x[1]
        $template = '$0$'
        if ([regex]::IsMatch($expr, '(?<!\\)\((?!\?)')) { $template = '$1$' }
        $kids.Add((JNode (JObj 'RegexExtractor' 'RegexExtractorGui' 'RegexExtractor' (Fill $N.api_regex @{ ID = $aid; VAR = $var }) $true @(
                        (PStr 'RegexExtractor.useHeaders' 'false'), (PStr 'RegexExtractor.refname' $var), (PStr 'RegexExtractor.regex' $expr),
                        (PStr 'RegexExtractor.template' $template), (PStr 'RegexExtractor.default' 'NOT_FOUND'),
                        (PBool 'RegexExtractor.default_empty_value' $false), (PStr 'RegexExtractor.match_number' '1')) (Fill $CM.api_regex @{ VAR = $var })) $null))
    }
    foreach ($x in $api.json) {
        $var = $x[0]
        $jp = $x[1]
        $kids.Add((JNode (JObj 'JSONPostProcessor' 'JSONPostProcessorGui' 'JSONPostProcessor' (Fill $N.api_json @{ ID = $aid; VAR = $var }) $true @(
                        (PStr 'JSONPostProcessor.referenceNames' $var), (PStr 'JSONPostProcessor.jsonPathExprs' $jp),
                        (PStr 'JSONPostProcessor.match_numbers' '1'), (PStr 'JSONPostProcessor.defaultValues' 'NOT_FOUND')) (Fill $CM.api_json @{ VAR = $var; PATH = $jp })) $null))
    }
    if ($api.post) { $kids.Add((JNode (Jsr223 'JSR223PostProcessor' (Fill $N.api_post $ids) $true '' $api.post $CM.api_post) $null)) }
    if ($api.statuses.Count -gt 0) {
        $codes = $api.statuses.ToArray()
        $ignore = $false
        foreach ($c in $codes) { if ([int]$c -ge 400) { $ignore = $true } }
        $strs = New-Object System.Collections.Generic.List[object]
        foreach ($c in $codes) { $strs.Add((PStr ([string](JavaHash $c)) $c)) }
        $tt = 8
        if ($codes.Length -gt 1) { $tt = 40 }
        $cmt = Fill $CM.api_status @{ CODES = ($codes -join ' / ') }
        if ($ignore) { $cmt += $CM.api_status_ignore }
        $kids.Add((JNode (JObj 'ResponseAssertion' 'AssertionGui' 'ResponseAssertion' (Fill $N.api_status @{ ID = $aid; CODES = ($codes -join ',') }) $true @(
                        (PColl 'Asserion.test_strings' $strs), (PStr 'Assertion.custom_message' ''), (PStr 'Assertion.test_field' 'Assertion.response_code'),
                        (PBool 'Assertion.assume_success' $ignore), (PInt 'Assertion.test_type' $tt)) $cmt) $null))
    }
    if ($api.texts.Count -gt 0) {
        $strs = New-Object System.Collections.Generic.List[object]
        foreach ($t in $api.texts) { $strs.Add((PStr ([string](JavaHash $t)) $t)) }
        $kids.Add((JNode (JObj 'ResponseAssertion' 'AssertionGui' 'ResponseAssertion' (Fill $N.api_text $ids) $true @(
                        (PColl 'Asserion.test_strings' $strs), (PStr 'Assertion.custom_message' ''), (PStr 'Assertion.test_field' 'Assertion.response_data'),
                        (PBool 'Assertion.assume_success' $false), (PInt 'Assertion.test_type' 16)) $CM.api_text) $null))
    }
    if ($api.max_ms -gt 0) {
        $ms = [string]$api.max_ms
        $kids.Add((JNode (JObj 'DurationAssertion' 'DurationAssertionGui' 'DurationAssertion' (Fill $N.api_duration @{ ID = $aid; MS = $ms }) $true @(, (PStr 'DurationAssertion.duration' $ms)) (Fill $CM.api_duration @{ MS = $ms })) $null))
    }
    return (JNode (JObj 'HTTPSamplerProxy' 'HttpTestSampleGui' 'HTTPSamplerProxy' $api.label $api.enabled $props $comment) $kids)
}
function Build-Jmx($cfg, $apis, $ctx) {
    $N = $script:Spec.names
    $CM = $script:Spec.comments
    $SC = $script:Spec.scripts
    $hosts = Assign-Hosts $cfg $apis
    $on = { param($k) return ([string]$cfg[$k] -cne 'OFF') }
    $enabled = { param($k) return ([string]$cfg[$k] -ceq 'ON') }
    $rd = PFun 'RESULT_DIR' $cfg['result_dir']
    $top = New-Object System.Collections.Generic.List[object]
    $args2 = New-Object System.Collections.Generic.List[object]
    $arg = {
        param($name, $value, $desc)
        return (PElem $name 'Argument' @((PStr 'Argument.name' $name), (PStr 'Argument.value' $value), (PStr 'Argument.desc' $desc), (PStr 'Argument.metadata' '=')) $null)
    }
    foreach ($h in $hosts) {
        $args2.Add((& $arg $h.pv (PFun $h.pv $h.scheme) (Fill $CM.udv_protocol @{ VAR = $h.pv })))
        $args2.Add((& $arg $h.hv (PFun $h.hv $h.host) (Fill $CM.udv_host @{ VAR = $h.hv })))
        $args2.Add((& $arg $h.portv (PFun $h.portv $h.port) (Fill $CM.udv_port @{ VAR = $h.portv })))
    }
    foreach ($uv in $cfg['_user_vars']) {
        $v = $uv[1]
        if (-not [regex]::IsMatch($v, '[${}()]')) { $v = PFun $uv[0] $uv[1] }
        $args2.Add((& $arg $uv[0] $v (Fill $CM.udv_user @{ VAR = $uv[0] })))
    }
    $top.Add((JNode (JObj 'Arguments' 'ArgumentsPanel' 'Arguments' $N.udv $true @(, (PColl 'Arguments.arguments' $args2)) $CM.udv) $null))
    if (& $on 'http_defaults') {
        $top.Add((JNode (JObj 'ConfigTestElement' 'HttpDefaultsGui' 'ConfigTestElement' $N.http_defaults (& $enabled 'http_defaults') @(
                        (PElem 'HTTPsampler.Arguments' 'Arguments' @(, (PColl 'Arguments.arguments' @())) @(@('guiclass', 'HTTPArgumentsPanel'), @('testclass', 'Arguments'), @('testname', 'User Defined Variables'), @('enabled', 'true'))),
                        (PStr 'HTTPSampler.domain' ''), (PStr 'HTTPSampler.port' ''), (PStr 'HTTPSampler.protocol' ''),
                        (PStr 'HTTPSampler.contentEncoding' $cfg['content_encoding']), (PStr 'HTTPSampler.path' ''),
                        (PStr 'HTTPSampler.concurrentPool' '6'),
                        (PStr 'HTTPSampler.connect_timeout' (PFun 'CONNECT_TIMEOUT' $cfg['connect_timeout'])),
                        (PStr 'HTTPSampler.response_timeout' (PFun 'RESPONSE_TIMEOUT' $cfg['response_timeout']))) $CM.http_defaults) $null))
    }
    if ($cfg['_common_headers'].Count -gt 0) { $top.Add((JNode (Header-Manager $N.common_headers $cfg['_common_headers'] $CM.common_headers) $null)) }
    if (& $on 'cookie_manager') {
        $domain = [string]$cfg['cookie_domain']
        if (-not $domain -and $hosts.Count -gt 0) { $domain = '${' + $hosts[0].hv + '}' }
        $cookies = New-Object System.Collections.Generic.List[object]
        foreach ($ck in $cfg['_cookies']) {
            $cookies.Add((XEl 'elementProp' @(@('name', $ck[0]), @('elementType', 'Cookie'), @('testname', $ck[0])) @(
                        (PStr 'Cookie.value' $ck[1]), (PStr 'Cookie.domain' $domain), (PStr 'Cookie.path' '/'), (PBool 'Cookie.secure' $false),
                        (PLong 'Cookie.expires' 0), (PBool 'Cookie.path_specified' $true), (PBool 'Cookie.domain_specified' $true)) $null))
        }
        $clear = 'false'
        if ($cfg['cookie_clear_each_iteration']) { $clear = 'true' }
        $top.Add((JNode (JObj 'CookieManager' 'CookiePanel' 'CookieManager' $N.cookie (& $enabled 'cookie_manager') @(
                        (PColl 'CookieManager.cookies' $cookies), (PBool 'CookieManager.clearEachIteration' $cfg['cookie_clear_each_iteration']),
                        (PBool 'CookieManager.controlledByThreadGroup' $false)) (Fill $CM.cookie @{ CLEAR = $clear })) $null))
    }
    $scheduler = ([string]$cfg['duration'] -cne '') -and ([long]$cfg['duration'] -gt 0)
    $setup = New-Object System.Collections.Generic.List[object]
    if (& $on 'setup_teardown') {
        $params = 'THREADS=' + (PFun 'THREADS' $cfg['threads']) + ' RAMP_UP=' + (PFun 'RAMP_UP' $cfg['ramp_up']) + ' LOOPS=' + (PFun 'LOOPS' $cfg['loops'])
        if ($scheduler) { $params += ' DURATION=' + (PFun 'DURATION' $cfg['duration']) }
        if ($hosts.Count -gt 0) { $params += ' HOST=${' + $hosts[0].hv + '}' }
        $params += ' RESULT_DIR=' + $rd
        $setup.Add((JNode (Jsr223 'JSR223Sampler' $N.setup_start (& $enabled 'setup_teardown') $params $SC.setup_start $CM.setup_start) $null))
    }
    if (& $on 'excel_data') {
        $params = (PFun 'EXCEL_FILE' $cfg['excel_file']) + '|' + (PFun 'EXCEL_SHEET' $cfg['excel_sheet'])
        $setup.Add((JNode (Jsr223 'JSR223Sampler' $N.excel_loader (& $enabled 'excel_data') $params $SC.excel_loader $CM.excel_loader) $null))
    }
    if ($setup.Count -gt 0) {
        $tgOn = $false
        foreach ($s in $setup) { if ($s.elem.attrs[3][1] -ceq 'true') { $tgOn = $true } }
        $top.Add((JNode (Thread-Group 'SetupThreadGroup' 'SetupThreadGroupGui' $N.setup_tg $tgOn '1' '1' '1' $false '' 'continue' $CM.setup_tg) $setup))
    }
    $main = New-Object System.Collections.Generic.List[object]
    if (& $on 'counter') {
        $scope = $CM.counter_scope_global
        if ($cfg['counter_per_thread']) { $scope = $CM.counter_scope_thread }
        $main.Add((JNode (JObj 'CounterConfig' 'CounterConfigGui' 'CounterConfig' (Fill $N.counter @{ NAME = $cfg['counter_name'] }) (& $enabled 'counter') @(
                        (PStr 'CounterConfig.start' $cfg['counter_start']), (PStr 'CounterConfig.end' ''), (PStr 'CounterConfig.incr' $cfg['counter_increment']),
                        (PStr 'CounterConfig.name' $cfg['counter_name']), (PStr 'CounterConfig.format' $cfg['counter_format']),
                        (PBool 'CounterConfig.per_user' $cfg['counter_per_thread'])) (Fill $CM.counter @{ NAME = $cfg['counter_name']; SCOPE = $scope })) $null))
    }
    if (& $on 'random_variable') {
        $main.Add((JNode (JObj 'RandomVariableConfig' 'TestBeanGUI' 'RandomVariableConfig' (Fill $N.random @{ NAME = $cfg['random_name'] }) (& $enabled 'random_variable') @(
                        (PStr 'maximumValue' $cfg['random_max']), (PStr 'minimumValue' $cfg['random_min']), (PStr 'outputFormat' ''),
                        (PBool 'perThread' $false), (PStr 'randomSeed' ''), (PStr 'variableName' $cfg['random_name'])) (Fill $CM.random @{ NAME = $cfg['random_name']; MIN = $cfg['random_min']; MAX = $cfg['random_max'] })) $null))
    }
    if (& $on 'csv_data') {
        $main.Add((JNode (JObj 'CSVDataSet' 'TestBeanGUI' 'CSVDataSet' $N.csv (& $enabled 'csv_data') @(
                        (PStr 'delimiter' $cfg['csv_delimiter']), (PStr 'fileEncoding' $cfg['csv_encoding']), (PStr 'filename' (PFun 'CSV_FILE' $cfg['csv_file'])),
                        (PBool 'ignoreFirstLine' $ctx.csv_skip_header), (PBool 'quotedData' $true), (PBool 'recycle' $cfg['csv_recycle']),
                        (PStr 'shareMode' ('shareMode.' + $cfg['csv_share_mode'])), (PBool 'stopThread' $cfg['csv_stop_thread']),
                        (PStr 'variableNames' $ctx.csv_vars)) (Fill $CM.csv @{ VARS = $ctx.csv_vars })) $null))
    }
    if (& $on 'excel_data') { $main.Add((JNode (Jsr223 'JSR223PreProcessor' $N.excel_row (& $enabled 'excel_data') '' $SC.excel_row $CM.excel_row) $null)) }
    if ([string]$cfg['access_token_dir'] -cne '') {
        $params = (PFun 'ACCESS_TOKEN_DIR' $cfg['access_token_dir']) + '|' + (PFun 'ACCESS_TOKEN_FILE' $cfg['access_token_file'])
        $main.Add((JNode (Jsr223 'JSR223PreProcessor' $N.access_token $true $params $SC.access_token $CM.access_token) $null))
    }
    if (& $on 'common_preprocessor') { $main.Add((JNode (Jsr223 'JSR223PreProcessor' $N.common_pre (& $enabled 'common_preprocessor') '' $SC.common_pre $CM.common_pre) $null)) }
    if (& $on 'common_postprocessor') { $main.Add((JNode (Jsr223 'JSR223PostProcessor' $N.common_post (& $enabled 'common_postprocessor') '' $SC.common_post $CM.common_post) $null)) }
    if (& $on 'throughput_timer') {
        $main.Add((JNode (JObj 'ConstantThroughputTimer' 'TestBeanGUI' 'ConstantThroughputTimer' $N.throughput (& $enabled 'throughput_timer') @(
                        (PInt 'calcMode' 4), (PStr 'throughput' (PFun 'THROUGHPUT' $cfg['throughput_per_min']))) (Fill $CM.throughput @{ VALUE = $cfg['throughput_per_min'] })) $null))
    }
    foreach ($a in $apis) { $main.Add((Api-Node $a $cfg)) }
    if (& $on 'debug_sampler') {
        $main.Add((JNode (JObj 'DebugSampler' 'TestBeanGUI' 'DebugSampler' $N.debug (& $enabled 'debug_sampler') @(
                        (PBool 'displayJMeterProperties' $false), (PBool 'displayJMeterVariables' $true), (PBool 'displaySystemProperties' $false)) $CM.debug) $null))
    }
    $dur = ''
    $durNote = ''
    if ($scheduler) { $dur = PFun 'DURATION' $cfg['duration']; $durNote = $CM.main_tg_duration }
    $top.Add((JNode (Thread-Group 'ThreadGroup' 'ThreadGroupGui' $N.main_tg $true (PFun 'THREADS' $cfg['threads']) (PFun 'RAMP_UP' $cfg['ramp_up']) (PFun 'LOOPS' $cfg['loops']) $scheduler $dur $cfg['on_sample_error'] (Fill $CM.main_tg @{ DURATION = $durNote })) $main))
    if (& $on 'setup_teardown') {
        $top.Add((JNode (Thread-Group 'PostThreadGroup' 'PostThreadGroupGui' $N.teardown_tg (& $enabled 'setup_teardown') '1' '1' '1' $false '' 'continue' $CM.teardown_tg) @(
                    , (JNode (Jsr223 'JSR223Sampler' $N.teardown_end (& $enabled 'setup_teardown') '' $SC.teardown_end $CM.teardown_end) $null))))
    }
    if (& $on 'view_results_tree') { $top.Add((JNode (Collector 'ViewResultsFullVisualizer' $N.view_tree (& $enabled 'view_results_tree') '' $CM.view_tree) $null)) }
    if (& $on 'summary_report') { $top.Add((JNode (Collector 'SummaryReport' $N.summary (& $enabled 'summary_report') '' $CM.summary) $null)) }
    if (& $on 'aggregate_report') { $top.Add((JNode (Collector 'StatVisualizer' $N.aggregate (& $enabled 'aggregate_report') '' $CM.aggregate) $null)) }
    if (& $on 'jtl_writer') { $top.Add((JNode (Collector 'SimpleDataWriter' $N.jtl (& $enabled 'jtl_writer') ($rd + '/result_all.jtl') (Fill $CM.jtl @{ DIR = $cfg['result_dir'] })) $null)) }
    if (& $on 'error_writer') { $top.Add((JNode (Collector 'SimpleDataWriter' $N.error_xml (& $enabled 'error_writer') ($rd + '/error_detail.xml') $CM.error_xml $true $true) $null)) }
    if (& $on 'result_log') {
        $params = $rd + '|' + (PFun 'LOG_RESPONSE' $cfg['log_response']) + '|' + (PFun 'LOG_MAX_CHARS' $cfg['log_max_chars'])
        $top.Add((JNode (Jsr223 'JSR223Listener' $N.result_log (& $enabled 'result_log') $params $SC.result_listener $CM.result_log) $null))
    }
    $plan = JObj 'TestPlan' 'TestPlanGui' 'TestPlan' $ctx.plan_name $true @(
        (PBool 'TestPlan.functional_mode' $false), (PBool 'TestPlan.tearDown_on_shutdown' $true), (PBool 'TestPlan.serialize_threadgroups' $false),
        (PElem 'TestPlan.user_defined_variables' 'Arguments' @(, (PColl 'Arguments.arguments' @())) @(@('guiclass', 'ArgumentsPanel'), @('testclass', 'Arguments'), @('testname', 'User Defined Variables'), @('enabled', 'true'))),
        (PStr 'TestPlan.user_define_classpath' '')) (Fill $CM.test_plan @{ VERSION = $script:ToolVersion; DATE = $ctx.date; INPUT = $ctx.input })
    return @((Jmx-Text (JNode $plan $top)), $hosts)
}

# ----------------------------------------------------------------------------- commands
function Setting-Default([string]$key) {
    foreach ($s in $script:Spec.settings) { if ($s.key -ceq $key) { return [string]$s.default } }
    return ''
}
function Resolve-Data([string]$path, [string]$outDir) {
    if ([System.IO.Path]::IsPathRooted($path)) { return $path }
    return (Join-Path $outDir $path)
}
function Prepare-Data($cfg, [string]$outDir, [bool]$force) {
    $created = New-Object System.Collections.Generic.List[string]
    foreach ($t in @(@('csv_data', 'csv_file', 'csv'), @('excel_data', 'excel_file', 'xlsx'))) {
        if ($cfg[$t[0]] -ceq 'OFF') { continue }
        $rel = [string]$cfg[$t[1]]
        $full = Resolve-Data $rel $outDir
        if ($rel.Replace('\', '/') -ceq (Setting-Default $t[1]) -and ($force -or -not (Test-Path -LiteralPath $full))) {
            if ($t[2] -ceq 'csv') { Build-TestdataCsv $full } else { Build-TestdataXlsx $full }
            $created.Add($full)
        } elseif (-not (Test-Path -LiteralPath $full)) {
            Write-Warn ('データファイルが見つかりません: ' + $full + '（JMX のフォルダからの相対パス。実行前に配置してください）')
        }
    }
    return ,$created
}
function Get-CsvVariables($cfg, [string]$outDir) {
    if ($cfg['csv_data'] -ceq 'OFF') { return @('', $true) }
    $delim = [string]$cfg['csv_delimiter']
    if ($delim -ceq '\t') { $delim = "`t" } elseif ($delim.Length -gt 0) { $delim = $delim.Substring(0, 1) } else { $delim = ',' }
    $full = Resolve-Data $cfg['csv_file'] $outDir
    $first = $null
    if (Test-Path -LiteralPath $full -PathType Leaf) {
        $tx = Read-TextAuto $full
        $rows = Parse-Csv $tx[0] $delim
        if ($rows.Count -gt 0) {
            $first = New-Object System.Collections.Generic.List[string]
            foreach ($c in $rows[0]) { $first.Add($c.Trim()) }
        }
    }
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($x in ([string]$cfg['csv_variables']).Split(',')) { if ($x.Trim()) { $names.Add($x.Trim()) } }
    if ($names.Count -gt 0) {
        $skip = $true
        if ($null -ne $first) {
            $skip = ($first.Count -eq $names.Count)
            if ($skip) { for ($k = 0; $k -lt $names.Count; $k++) { if ($first[$k].ToLowerInvariant() -cne $names[$k].ToLowerInvariant()) { $skip = $false } } }
        }
        return @(($names.ToArray() -join ','), $skip)
    }
    if ($null -ne $first) {
        $any = $false
        $all = $true
        foreach ($c in $first) { if ($c) { $any = $true } else { $all = $false } }
        if ($any) {
            if (-not $all) { Write-Warn ('CSV の 1 行目に空の列名があります: ' + $full) }
            return @(($first.ToArray() -join ','), $true)
        }
    }
    Write-Warn 'CSV の変数名を決められないため、サンプルの列名を使用します（設定 csv_variables で指定してください）'
    return @((@($script:Spec.sample_csv.header) -join ','), $true)
}
function Safe-FileName([string]$name, [string]$default) {
    $base = [System.IO.Path]::GetFileName($name.Replace('\', '/').Replace('/', [System.IO.Path]::DirectorySeparatorChar))
    $s = [regex]::Replace($base, '[^A-Za-z0-9._-]', '_').Trim([char[]]@('.', '_'))
    if (-not $s) { return $default }
    return $s
}
function Write-Jmx([string]$path, [string]$text) {
    if (Test-Path -LiteralPath $path) {
        $bak = $path + '.' + (Get-Stamp) + '.bak'
        Move-Item -LiteralPath $path -Destination $bak
        Write-Info ('既存の JMX を退避しました: ' + $bak)
    }
    Write-TextFile $path $text
}
function Invoke-Generate {
    if (-not $InputFile) { throw (New-ToolError '入力ファイルを -InputFile（-i）で指定してください（例: -i input\api_definition.xlsx）') }
    if (-not (Test-Path -LiteralPath $InputFile -PathType Leaf)) { throw (New-ToolError ('入力ファイルが見つかりません: ' + $InputFile)) }
    $out = $OutDir
    if (-not $out) { $out = 'output' }
    Write-Info ('入力ファイル: ' + $InputFile)
    $loaded = Load-Input $InputFile $SettingsFile
    $cfg = Resolve-Settings $loaded[0]
    if ($JmxName) { $cfg['output_file'] = $JmxName }
    $apis = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($rec in $loaded[1]) {
        $a = Normalize-Api $rec $i $cfg
        if ($null -ne $a) { $apis.Add($a) }
        $i++
    }
    if ($apis.Count -eq 0) { throw (New-ToolError '有効な API 定義が 1 件もありません') }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($a in $apis) { if (-not $seen.Add($a.id)) { Write-Warn ('API の番号（No）が重複しています: ' + $a.id) } }
    if (-not (Test-Path -LiteralPath $out)) { [void](New-Item -ItemType Directory -Path $out) }
    foreach ($f in (Prepare-Data $cfg $out ([bool]$Force))) { Write-Info ('サンプルデータを作成しました: ' + $f) }
    $cv = Get-CsvVariables $cfg $out
    $outName = Safe-FileName $cfg['output_file'] 'api_test.jmx'
    if (-not $outName.ToLowerInvariant().EndsWith('.jmx', $script:Ord)) { $outName += '.jmx' }
    if ($outName -cne $cfg['output_file']) { Write-Info ('出力ファイル名: ' + $outName + '（指定値「' + $cfg['output_file'] + '」から変換）') }
    $ctx = @{ input = [System.IO.Path]::GetFileName($InputFile); date = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $script:Inv)
        csv_vars = [string]$cv[0]; csv_skip_header = [bool]$cv[1]; plan_name = [string]$cfg['test_plan_name'] }
    $outputs = New-Object System.Collections.Generic.List[string]
    if ($Split) {
        $stem = $outName.Substring(0, $outName.Length - 4)
        $used = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($a in $apis) {
            $c = $ctx.Clone()
            $c.plan_name = [string]$cfg['test_plan_name'] + ' - ' + $a.label
            $res = Build-Jmx $cfg @(, $a) $c
            $fname = $stem + '_' + $a.id
            $k = 2
            while ($used.Contains($fname.ToLowerInvariant())) { $fname = $stem + '_' + $a.id + '_' + $k; $k++ }
            [void]$used.Add($fname.ToLowerInvariant())
            $path = Join-Path $out ($fname + '.jmx')
            Write-Jmx $path $res[0]
            $outputs.Add($path)
        }
    } else {
        $res = Build-Jmx $cfg $apis $ctx
        $path = Join-Path $out $outName
        Write-Jmx $path $res[0]
        $outputs.Add($path)
    }
    $nOn = 0
    foreach ($a in $apis) { if ($a.enabled) { $nOn++ } }
    Write-Info ('API: ' + $apis.Count + ' 件（有効 ' + $nOn + ' / 無効 ' + ($apis.Count - $nOn) + '）')
    foreach ($p in $outputs) { Write-Info ('JMX を生成しました: ' + $p) }
    Write-Host ''
    Write-Host '次の手順:'
    Write-Host ('  GUI で確認 : %JMETER_HOME%\bin\jmeter.bat で開く（File → Open）: ' + $outputs[0])
    Write-Host ('  CLI で実行 : .\jmeter_api_tool.bat run -t "' + $outputs[0] + '"')
    if ($script:Warnings.Count -gt 0) { Write-Host ''; Write-Host ('警告が ' + $script:Warnings.Count + ' 件あります（[WARN] を確認してください）') -ForegroundColor Yellow }
    return 0
}
function Write-IfAbsent([string]$path, [bool]$force, [scriptblock]$builder) {
    if ((Test-Path -LiteralPath $path) -and -not $force) { Write-Warn ('既に存在するため作成しません（上書きは -Force）: ' + $path); return }
    & $builder $path
    Write-Info ('作成しました: ' + $path)
}
function Invoke-Init {
    $out = $OutDir
    if (-not $out) { $out = 'input' }
    if (-not (Test-Path -LiteralPath $out)) { [void](New-Item -ItemType Directory -Path $out) }
    Write-IfAbsent (Join-Path $out 'api_definition.xlsx') ([bool]$Force) { param($p) Build-TemplateXlsx $p }
    Write-IfAbsent (Join-Path $out 'api_definition.csv') ([bool]$Force) { param($p) Write-TextFile $p (To-Csv (Template-ApiRows)) $true }
    Write-IfAbsent (Join-Path $out 'api_definition_settings.csv') ([bool]$Force) { param($p) Write-TextFile $p (To-Csv (Template-SettingsRows)) $true }
    Write-IfAbsent (Join-Path $out 'openapi_config.csv') ([bool]$Force) { param($p) Write-TextFile $p (To-Csv (Template-SettingsRows $script:Spec.openapi_config)) $true }
    Write-Host ''
    Write-Host '次の手順: テンプレートに API を記入し、generate コマンドで JMX を生成します。'
    Write-Host ('  .\jmeter_api_tool.bat generate -i ' + (Join-Path $out 'api_definition.xlsx'))
    Write-Host ('  .\jmeter_api_tool.bat generate -i openapi.json -c ' + (Join-Path $out 'openapi_config.csv') + '   # OpenAPI 定義から生成する場合')
    return 0
}
function Invoke-Manual {
    $out = $OutDir
    if (-not $out) { $out = 'docs' }
    if (-not (Test-Path -LiteralPath $out)) { [void](New-Item -ItemType Directory -Path $out) }
    foreach ($p in (Build-Manual $out)) { Write-Info ('作成しました: ' + $p) }
    return 0
}

# ----------------------------------------------------------------------------- run
function Find-JMeterHome([string]$explicit) {
    if ($explicit) {
        if (Test-Path -LiteralPath (Join-Path $explicit 'bin\ApacheJMeter.jar')) { return (Resolve-Path -LiteralPath $explicit).Path }
        throw (New-ToolError ('指定されたフォルダに JMeter がありません: ' + $explicit))
    }
    $cands = New-Object System.Collections.Generic.List[string]
    if ($env:JMETER_HOME) { $cands.Add($env:JMETER_HOME) }
    $cmd = Get-Command 'jmeter.bat' -ErrorAction SilentlyContinue
    if ($null -ne $cmd) { $cands.Add((Split-Path -Parent (Split-Path -Parent $cmd.Source))) }
    foreach ($x in @('C:\apache-jmeter-5.6.3', 'C:\tools\apache-jmeter-5.6.3', 'C:\jmeter\apache-jmeter-5.6.3', 'D:\apache-jmeter-5.6.3')) { $cands.Add($x) }
    if ($env:USERPROFILE) { $cands.Add((Join-Path $env:USERPROFILE 'apache-jmeter-5.6.3')) }
    if ($env:ProgramFiles) { $cands.Add((Join-Path $env:ProgramFiles 'apache-jmeter-5.6.3')) }
    if ($PSScriptRoot) { $cands.Add((Join-Path $PSScriptRoot 'apache-jmeter-5.6.3')) }
    foreach ($c in $cands) {
        if ($c -and (Test-Path -LiteralPath (Join-Path $c 'bin\ApacheJMeter.jar'))) { return (Resolve-Path -LiteralPath $c).Path }
    }
    return $null
}
function Find-Java {
    if ($env:JAVA_HOME -and (Test-Path -LiteralPath (Join-Path $env:JAVA_HOME 'bin\java.exe'))) { return (Join-Path $env:JAVA_HOME 'bin\java.exe') }
    $cmd = Get-Command 'java.exe' -ErrorAction SilentlyContinue
    if ($null -ne $cmd) { return $cmd.Source }
    return $null
}
function Get-JavaVersion([string]$java) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $java
    $psi.Arguments = '-version'
    $psi.UseShellExecute = $false
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true
    $psi.CreateNoWindow = $true
    $p = [System.Diagnostics.Process]::Start($psi)
    $text = $p.StandardError.ReadToEnd() + $p.StandardOutput.ReadToEnd()
    $p.WaitForExit()
    $m = [regex]::Match($text, 'version "([^"]+)"')
    if ($m.Success) { return $m.Groups[1].Value }
    return ''
}
function Get-JMeterVersion([string]$jar) {
    try {
        $fs = New-Object System.IO.FileStream($jar, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $e = $zip.GetEntry('META-INF/MANIFEST.MF')
            if ($null -eq $e) { return '' }
            $sr = New-Object System.IO.StreamReader($e.Open())
            $mf = $sr.ReadToEnd()
            $sr.Dispose()
            $m = [regex]::Match($mf, 'Implementation-Version:\s*(\S+)')
            if ($m.Success) { return $m.Groups[1].Value }
            return ''
        } finally { $zip.Dispose(); $fs.Dispose() }
    } catch { return '' }
}
function Disp-Width([string]$s) {
    $w = 0
    foreach ($ch in $s.ToCharArray()) { $c = [int]$ch; if ($c -ge 0x1100 -and -not ($c -ge 0xFF61 -and $c -le 0xFF9F)) { $w += 2 } else { $w += 1 } }
    return $w
}
function Pad-Disp([string]$s, [int]$width) { return $s + (' ' * [Math]::Max(0, $width - (Disp-Width $s))) }
function RPad-Disp([string]$s, [int]$width) { return (' ' * [Math]::Max(0, $width - (Disp-Width $s))) + $s }
function Show-Summary([string]$path) {
    $tx = Read-TextAuto $path
    $stats = $tx[0] | ConvertFrom-Json
    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($p in $stats.PSObject.Properties) { if ($p.Name -cne 'Total') { $keys.Add($p.Name) } }
    $keys.Sort([System.StringComparer]::Ordinal)
    if (Has-Prop $stats 'Total') { $keys.Add('Total') }
    $heads = New-Object System.Collections.Generic.List[string]
    foreach ($hw in @(@('件数', 8), @('エラー', 7), @('エラー率', 9), @('平均ms', 9), @('90%ms', 9), @('最大ms', 9), @('件/秒', 10))) { $heads.Add((RPad-Disp $hw[0] $hw[1])) }
    Write-Host ((Pad-Disp 'ラベル' 36) + ($heads.ToArray() -join ' '))
    foreach ($k in $keys) {
        $s = $stats.PSObject.Properties[$k].Value
        $label = [string](Get-Prop $s 'transaction' $k)
        if ($label.Length -gt 34) { $label = $label.Substring(0, 34) }
        Write-Host ((Pad-Disp $label 36) + [string]::Format($script:Inv, '{0,8} {1,7} {2,8:F2}% {3,9:F1} {4,9:F1} {5,9:F1} {6,10:F2}',
                [long](Get-Prop $s 'sampleCount' 0), [long](Get-Prop $s 'errorCount' 0), [double](Get-Prop $s 'errorPct' 0),
                [double](Get-Prop $s 'meanResTime' 0), [double](Get-Prop $s 'pct1ResTime' 0), [double](Get-Prop $s 'maxResTime' 0),
                [double](Get-Prop $s 'throughput' 0)))
    }
    if ((Has-Prop $stats 'Total') -and [long](Get-Prop $stats.Total 'errorCount' 0) -gt 0) { return 3 }
    return 0
}
function Invoke-Run {
    if (-not $Jmx) { throw (New-ToolError '実行する JMX を -Jmx（-t）で指定してください（例: -t output\api_test.jmx）') }
    if (-not (Test-Path -LiteralPath $Jmx -PathType Leaf)) { throw (New-ToolError ('JMX ファイルが見つかりません: ' + $Jmx)) }
    foreach ($v in @($Threads, $RampUp, $Loops, $Duration)) {
        if ($v -and -not [regex]::IsMatch($v, '^-?[0-9]+$')) { throw (New-ToolError ('数値を指定してください: ' + $v)) }
    }
    $props2 = New-Object System.Collections.Generic.List[string]
    foreach ($p in $Prop) { foreach ($x in ([string]$p).Split(',')) { if ($x.Trim()) { $props2.Add($x.Trim()) } } }
    foreach ($p in $props2) { if (-not [regex]::IsMatch($p, '^[^=]+=')) { throw (New-ToolError ('-Prop は 名前=値 の形式で指定してください: ' + $p)) } }
    if ($Heap -and -not [regex]::IsMatch($Heap, '^[0-9]+[mMgG]$')) { throw (New-ToolError ('-Heap は 512m / 2g の形式で指定してください: ' + $Heap)) }
    $jhome = Find-JMeterHome $JMeterHome
    if (-not $jhome) { throw (New-ToolError 'JMeter が見つかりません。環境変数 JMETER_HOME を設定するか -JMeterHome で指定してください。') }
    Write-Info ('JMeter: ' + $jhome)
    $jver = Get-JMeterVersion (Join-Path $jhome 'bin\ApacheJMeter.jar')
    if ($jver) {
        Write-Info ('JMeter バージョン: ' + $jver)
        if (-not $jver.StartsWith('5.6', $script:Ord)) { Write-Warn ('このツールは JMeter 5.6.3 を想定しています（検出: ' + $jver + '）') }
    }
    $java = Find-Java
    if (-not $java) { throw (New-ToolError 'Java が見つかりません。JDK 17 をインストールし、JAVA_HOME または PATH を設定してください。') }
    $ver = Get-JavaVersion $java
    $major = 0
    if ($ver) {
        $parts = $ver.Split('.')
        $first = 0
        [void][int]::TryParse(([regex]::Match($parts[0], '^[0-9]+')).Value, [ref]$first)
        $major = $first
        if ($first -eq 1 -and $parts.Length -gt 1) { [void][int]::TryParse(([regex]::Match($parts[1], '^[0-9]+')).Value, [ref]$major) }
        if ($major -gt 0 -and $major -lt 8) { throw (New-ToolError ('Java 8 以上が必要です（現在: ' + $ver + '）。Java 17 を推奨します。')) }
        Write-Info ('Java: ' + $ver + ' (' + $java + ')')
    } else {
        Write-Warn ('Java のバージョンを判別できませんでした: ' + $java)
    }
    $outBase = $OutDir
    if (-not $outBase) { $outBase = 'result' }
    if (-not (Test-Path -LiteralPath $outBase)) { [void](New-Item -ItemType Directory -Path $outBase) }
    $name = [System.IO.Path]::GetFileNameWithoutExtension($Jmx)
    $runDir = Join-Path (Resolve-Path -LiteralPath $outBase).Path ($name + '_' + (Get-Stamp))
    [void](New-Item -ItemType Directory -Path $runDir)
    $jmxAbs = (Resolve-Path -LiteralPath $Jmx).Path
    $jvm = New-Object System.Collections.Generic.List[string]
    if ($major -ge 9 -or $major -eq 0) {
        foreach ($x in @('java.desktop/sun.awt', 'java.desktop/sun.swing', 'java.desktop/javax.swing.text.html', 'java.desktop/java.awt',
                'java.desktop/java.awt.font', 'java.base/java.lang', 'java.base/java.lang.invoke', 'java.base/java.lang.reflect',
                'java.base/java.util', 'java.base/java.text', 'java.desktop/sun.awt.shell')) { $jvm.Add('--add-opens=' + $x + '=ALL-UNNAMED') }
    }
    $jvm.Add('-XX:+HeapDumpOnOutOfMemoryError')
    if ($Heap) { $h = $Heap } else { $h = '1g' }
    foreach ($x in @(('-Xms' + $h), ('-Xmx' + $h), '-XX:MaxMetaspaceSize=256m', '-XX:+UseG1GC', '-XX:MaxGCPauseMillis=100', '-XX:G1ReservePercent=20',
            '-Djava.security.egd=file:/dev/urandom', '-Duser.language=en', '-Duser.region=EN')) { $jvm.Add($x) }
    if ($env:JVM_ARGS) { foreach ($x in ($env:JVM_ARGS -split '\s+')) { if ($x) { $jvm.Add($x) } } }
    $jargs = New-Object System.Collections.Generic.List[string]
    foreach ($x in @('-n', '-t', $jmxAbs, '-l', (Join-Path $runDir 'result.jtl'), '-j', (Join-Path $runDir 'jmeter.log'), '-e', '-o', (Join-Path $runDir 'report'),
            ('-JRESULT_DIR=' + $runDir), '-Jjmeter.save.saveservice.subresults=false')) { $jargs.Add($x) }
    if ($Threads) { $jargs.Add('-JTHREADS=' + $Threads) }
    if ($RampUp) { $jargs.Add('-JRAMP_UP=' + $RampUp) }
    if ($Loops) { $jargs.Add('-JLOOPS=' + $Loops) }
    if ($Duration) { $jargs.Add('-JDURATION=' + $Duration) }
    foreach ($p in $props2) { $jargs.Add('-J' + $p) }
    if ($ExtraArgs) { foreach ($x in ($ExtraArgs.Trim() -split '\s+')) { if ($x) { $jargs.Add($x) } } }
    Write-Info ('ヒープ: -Xms' + $h + ' -Xmx' + $h)
    Write-Info ('結果フォルダ: ' + $runDir)
    Write-Info ('実行コマンド: jmeter ' + ($jargs.ToArray() -join ' '))
    Write-Host ''
    $all = New-Object System.Collections.Generic.List[string]
    foreach ($x in $jvm) { $all.Add($x) }
    $all.Add('-jar')
    $all.Add((Join-Path $jhome 'bin\ApacheJMeter.jar'))
    foreach ($x in $jargs) { $all.Add($x) }
    $allArr = $all.ToArray()
    & $java @allArr | Out-Host
    $rc = $LASTEXITCODE
    Write-Host ''
    $summaryRc = 0
    $stat = Join-Path $runDir 'report\statistics.json'
    $jtl = Join-Path $runDir 'result.jtl'
    if (Test-Path -LiteralPath $stat) {
        Write-Host '==== 集計結果（API ごと） ===='
        $summaryRc = Show-Summary $stat
        Write-Host ''
    } elseif (Test-Path -LiteralPath $jtl) {
        $lines = @(Get-Content -LiteralPath $jtl -Encoding UTF8)
        $errors = 0
        for ($k = 1; $k -lt $lines.Count; $k++) { $f = $lines[$k].Split(','); if ($f.Length -gt 7 -and $f[7] -ceq 'false') { $errors++ } }
        Write-Host ('==== 集計結果 ==== 件数: ' + ($lines.Count - 1) + ' / エラー（概算）: ' + $errors)
        if ($errors -gt 0) { $summaryRc = 3 }
    }
    Write-Host ('結果フォルダ : ' + $runDir)
    foreach ($pair in @(@('report\index.html', 'HTML レポート'), @('result.jtl', 'JTL          '), @('jmeter.log', 'JMeter ログ  '), @('api_result.log', '結果ログ     '))) {
        $f = Join-Path $runDir $pair[0]
        if (Test-Path -LiteralPath $f) { Write-Host ('  ' + $pair[1] + ' : ' + $f) }
    }
    $ex = Join-Path $runDir 'error_detail.xml'
    if ((Test-Path -LiteralPath $ex) -and (Select-String -LiteralPath $ex -Pattern '<httpSample|<sample' -Quiet)) { Write-Host ('  エラー詳細    : ' + $ex) }
    if ($rc -ne 0) {
        Write-Host ('[ERROR] JMeter が異常終了しました（終了コード ' + $rc + '）。jmeter.log の末尾:') -ForegroundColor Red
        $log = Join-Path $runDir 'jmeter.log'
        if (Test-Path -LiteralPath $log) { Get-Content -LiteralPath $log -Tail 20 | ForEach-Object { Write-Host $_ } }
        return $rc
    }
    if ($summaryRc -eq 3) {
        Write-Warn 'エラーのリクエストがあります。HTML レポート・結果ログを確認してください。'
        if ($FailOnError) { return 2 }
    }
    return 0
}

function Show-Usage {
    Write-Host @'
jmeter_api_tool.ps1 - REST API の JMeter テスト（JMX）生成・実行ツール
（.ps1 の実行が禁止されている環境では jmeter_api_tool.bat から実行してください）

使い方:
  jmeter_api_tool.bat init     [-OutDir フォルダ] [-Force]
      入力テンプレート（api_definition.xlsx / .csv / _settings.csv）を作成します（既定: input）

  jmeter_api_tool.bat generate -InputFile 入力ファイル [-SettingsFile 設定CSV] [-OutDir フォルダ] [-JmxName ファイル名] [-Split] [-Force] [OpenAPI用オプション]
      Excel（.xlsx）・CSV・OpenAPI 定義（.json）の API 定義から JMX を生成します（既定: output）
        -InputFile（-i）     入力ファイル（.xlsx / .csv / .json）
        -SettingsFile（-s）  設定 CSV（省略時は「入力ファイル名_settings.csv」→ 既定値）
        -OutDir（-o）        出力先フォルダ
        -JmxName（-n）       出力する JMX のファイル名（設定 output_file より優先。例: -n order_test.jmx）
        -Split               API ごとに別々の JMX を出力
        -Force（-f）         サンプルデータ（data\testdata.*）を上書き
      OpenAPI 定義（.json）用のオプション:
        -OpenApiConfig（-c） OpenAPI 用設定ファイル（接続先 PROTOCOL / HOST / PORT・HTTP ヘッダ・アクセストークンの
                             フォルダ TOKEN_DIR・定義の変数 PUBLIC_BASE_URL / NAME_PREFIX。
                             テンプレート: init で作成される input\openapi_config.csv）
        -Seed N              テストデータの乱数シード（同じ値なら同じデータ。既定: 毎回ランダム）
        -AnyMethod M         x-amazon-apigateway-any-method（ANY）で送るメソッド（既定: GET）
        -RequiredOnly        必須（required）のパラメータ・プロパティだけを生成
        -UseExamples         example / default があればランダム値より優先
        -PlaceholderValue N  値の位置に引用符なしで書かれた ${...} に仮に入れる数値（既定: 3000）

  jmeter_api_tool.bat manual   [-OutDir フォルダ]
      利用方法（JMeter_Usage_Guide.xlsx / .md）を出力します（既定: docs）

  jmeter_api_tool.bat run -Jmx JMX [オプション]
      JMeter を CLI モードで実行し、result\JMX名_日時\ に JTL・ログ・HTML レポートを出力します
        -Jmx（-t）           実行する JMX
        -Threads N           スレッド数（-JTHREADS）
        -RampUp SEC          ランプアップ秒数（-JRAMP_UP）
        -Loops N             ループ回数（-JLOOPS、-1 = 無限）
        -Duration SEC        実行時間（-JDURATION、duration 指定で生成した JMX のみ）
        -Prop 名前=値,...    任意のプロパティ（カンマ区切りで複数。例: -Prop HOST=stg.example.com）
        -OutDir（-o）        結果の出力先（既定: result）
        -Heap SIZE           JVM ヒープ（例: 2g。既定: 1g）
        -JMeterHome DIR      JMeter のフォルダ（省略時: JMETER_HOME → PATH → C:\apache-jmeter-5.6.3 等を探索）
        -FailOnError         エラーが 1 件でもあれば終了コード 2
        -ExtraArgs "..."     JMeter に渡す追加引数（例: "-H proxy.example.com -P 8080"）

  jmeter_api_tool.bat help | version

例:
  jmeter_api_tool.bat init
  jmeter_api_tool.bat generate -i input\api_definition.xlsx
  jmeter_api_tool.bat generate -i openapi.json -Seed 12345
  jmeter_api_tool.bat generate -i openapi.json -c input\openapi_config.csv
  jmeter_api_tool.bat run -t output\api_test.jmx -Threads 10 -Loops 100
'@
}

# ----------------------------------------------------------------------------- main
$exitCode = 0
try {
    switch ($Command.ToLowerInvariant()) {
        'init' { $exitCode = Invoke-Init }
        'generate' { $exitCode = Invoke-Generate }
        'manual' { $exitCode = Invoke-Manual }
        'run' { $exitCode = Invoke-Run }
        'help' { Show-Usage }
        '-h' { Show-Usage }
        'version' { Write-Host ('jmeter_api_tool ' + $script:ToolVersion) }
        default { Write-Host ('[ERROR] 不明なコマンドです: ' + $Command) -ForegroundColor Red; Show-Usage; $exitCode = 1 }
    }
} catch [System.InvalidOperationException] {
    Write-Host ('[ERROR] ' + $_.Exception.Message) -ForegroundColor Red
    $exitCode = 1
} catch {
    Write-Host ('[ERROR] ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host ('        (' + $_.InvocationInfo.PositionMessage.Trim() + ')') -ForegroundColor DarkGray
    $exitCode = 1
}
exit $exitCode
