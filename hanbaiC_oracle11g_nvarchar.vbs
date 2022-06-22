' ************************************************
' 管理者権限でコマンドプロンプトで実行を強制して、
' 最後に pause する
' ************************************************
Set Shell = CreateObject("Shell.Application")
Set WshShell = Wscript.CreateObject("WScript.Shell")
if Wscript.Arguments.Count = 0 then
	ScriptFullName = WScript.ScriptFullName
	Shell.ShellExecute "cmd.exe", "/c cscript.exe """ & ScriptFullName & """ dummy_param & pause", "", "runas", 1
	WScript.Quit
end if

' ************************************************
' 基本設定
' Microsoft ODBC for Oracle で実行できます
' DSN を作成して動作確認して指定して下さい
' ( 参考:http://lightbox.matrix.jp/ginpro/patio.cgi?mode=view&no=225&type=ref )
' ************************************************
' このスクリプトが存在するディレクトリを取得
strCurDir = WScript.ScriptFullName
strCurDir = Replace( strCurDir, WScript.ScriptName, "" )
strMdbPath = strCurDir & "販売管理C.mdb"

' Oracle のホスト文字列
' ( ローカル・ネット・サービス名 )
strTarget = "{Oracle in instantclient_21_6}"	' ODBC ドライバ
strDBQ = "localhost:1521/XEPDB1"	' ネット・サービス名として XE のみでも OK
' スキーマ(ユーザ)
strSc = "LIGHTBOX02"
' パスワード
strPwd = InputBox("パスワードを入力して下さい")

strDummy = "DUMMY" & Replace(Date,"/","") & Replace(Time,":","")

strMessage = "対象 MDB は " & strMdbPath & "です" & vbCrLf & vbCrLf

strMessage = strMessage & "▼ Oracleの環境です" & vbCrLf
strMessage = strMessage & "ODBC ドライバ : " & strTarget & vbCrLf
strMessage = strMessage & "インスタンス : " & strDBQ & vbCrLf
strMessage = strMessage & "USER(スキーマ) : " & strSc & vbCrLf
strMessage = strMessage & "PASS : " & strPwd & vbCrLf & vbCrLf
strMessage = strMessage & "一時テーブル : " & strDummy & vbCrLf & vbCrLf

strMessage = strMessage & "既にテーブルが存在する場合はメッセージが出ません" & vbCrLf
strMessage = strMessage & "それ以外ではエラーメッセージが出ますが、問題ありません"
if vbCancel = MsgBox( strMessage, vbOkCancel ) then
	Wscript.Quit
end if

' ************************************************
' 処理用文字列設定
' ************************************************
' MDB の接続文字列
strConnectMdb = _
"Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & strMdbPath & ";"

' Microsoft の Oracle 用 ODBC ドライバの接続文字列 (1)
strConnectOracle = _
"[ODBC;Driver=" & strTarget & ";DBQ=" & strDBQ &";UID=" & strSc & ";PWD=" & strPwd & "]"

' Microsoft の Oracle 用 ODBC ドライバの接続文字列 (2)
strConnectOracle2 = _
"Provider=MSDASQL;Driver=" & strTarget & ";DBQ=" & strDBQ &";UID=" & strSc & ";PWD=" & strPwd

' ************************************************
' 初期処理
' ************************************************
Set Cn = CreateObject("ADODB.Connection")
Set Cn2 = CreateObject("ADODB.Connection")
Cn.Open strConnectMdb
Cn2.Open strConnectOracle2

' ************************************************
' コード名称マスタ
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	区分 NUMBER(4,0)" & _
"	,コード NVARCHAR2(10)" & _
"	,名称 NVARCHAR2(50)" & _
"	,数値1 NUMBER(8,0)" & _
"	,数値2 NUMBER" & _
"	,作成日 DATE" & _
"	,更新日 DATE" & _
"	,primary key(区分,コード)" & _
")"
Call OracleTransfer( "コード名称マスタ", "[コード名称マスタ]", Query )

' ************************************************
' コントロールマスタ
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	キー NVARCHAR2(1)" & _
"	,売上日付 DATE" & _
"	,売上伝票 NUMBER(8,0)" & _
"	,会社名 NVARCHAR2(50)" & _
"	,組織コード NVARCHAR2(4)" & _
"	,起算月 NUMBER(2,0)" & _
"	,primary key(キー)" & _
")"
Call OracleTransfer( "コントロールマスタ","[コントロールマスタ]", Query )

' ************************************************
' メッセージマスタ
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	コード NVARCHAR2(4)" & _
"	,メッセージ NVARCHAR2(100)" & _
"	,primary key(コード)" & _
")"
Call OracleTransfer( "メッセージマスタ","[メッセージマスタ]", Query )

' ************************************************
' 取引データ
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	取引区分 NVARCHAR2(2)" & _
"	,伝票番号 NUMBER(8,0)" & _
"	,行 NUMBER(2,0)" & _
"	,取引日付 DATE" & _
"	,取引先コード NVARCHAR2(4)" & _
"	,商品コード NVARCHAR2(4)" & _
"	,数量 NUMBER" & _
"	,単価 NUMBER" & _
"	,金額 NUMBER" & _
"	,更新済 NVARCHAR2(1)" & _
"	,primary key(取引区分,伝票番号,行)" & _
")"
Call OracleTransfer( "取引データ","[取引データ]", Query )

' ************************************************
' 商品マスタ
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	商品コード NVARCHAR2(4)" & _
"	,商品名 NVARCHAR2(50)" & _
"	,在庫評価単価 NUMBER" & _
"	,販売単価 NUMBER" & _
"	,商品分類 NVARCHAR2(3)" & _
"	,商品区分 NVARCHAR2(1)" & _
"	,作成日 DATE" & _
"	,更新日 DATE" & _
"	,備考 NVARCHAR2(2000)" & _
"	,削除フラグ NVARCHAR2(1)" & _
"	,primary key(商品コード)" & _
")"
Call OracleTransfer( "商品マスタ","[商品マスタ]", Query )

' ************************************************
' 商品分類マスタ
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	商品分類 NVARCHAR2(3)" & _
"	,名称 NVARCHAR2(50)" & _
"	,作成日 DATE" & _
"	,更新日 DATE" & _
"	,primary key(商品分類)" & _
")"
Call OracleTransfer( "商品分類マスタ","[商品分類マスタ]", Query )

' ************************************************
' 得意先マスタ
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	得意先コード NVARCHAR2(4)" & _
"	,得意先名 NVARCHAR2(50)" & _
"	,得意先区分 NVARCHAR2(1)" & _
"	,担当者 NVARCHAR2(4)" & _
"	,郵便番号 NVARCHAR2(7)" & _
"	,住所１ NVARCHAR2(100)" & _
"	,住所２ NVARCHAR2(100)" & _
"	,作成日 DATE" & _
"	,更新日 DATE" & _
"	,締日 NUMBER(2,0)" & _
"	,締日区分 NUMBER(1,0)" & _
"	,支払日 NUMBER(2,0)" & _
"	,備考 NVARCHAR2(100)" & _
"	,primary key(得意先コード)" & _
")"
Call OracleTransfer( "得意先マスタ","[得意先マスタ]", Query )

' ************************************************
' 社員マスタ
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	社員コード NVARCHAR2(4)" & _
"	,氏名 NVARCHAR2(50)" & _
"	,フリガナ NVARCHAR2(50)" & _
"	,所属 NVARCHAR2(4)" & _
"	,性別 NUMBER(1,0)" & _
"	,作成日 DATE" & _
"	,更新日 DATE" & _
"	,給与 NUMBER" & _
"	,手当 NUMBER" & _
"	,管理者 NVARCHAR2(4)" & _
"	,生年月日 DATE" & _
"	,primary key(社員コード)" & _
")"
Call OracleTransfer( "社員マスタ","[社員マスタ]", Query )

' ************************************************
' 郵便番号マスタ
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	郵便番号 NVARCHAR2(7)" & _
"	,都道府県名カナ NVARCHAR2(255)" & _
"	,市区町村名カナ NVARCHAR2(255)" & _
"	,町域名カナ NVARCHAR2(255)" & _
"	,都道府県名 NVARCHAR2(255)" & _
"	,市区町村名 NVARCHAR2(255)" & _
"	,町域名 NVARCHAR2(255)" & _
")"
Call OracleTransfer( "郵便番号マスタ","[郵便番号マスタ]", Query )

' ************************************************
' 入金予定データ
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	得意先コード NVARCHAR2(4)" & _
"	,支払日 DATE" & _
"	,伝票合計金額 NUMBER(10,0)" & _
"	,伝票番号 NUMBER(10,0)" & _
")"
Call OracleTransfer( "入金予定データ","[入金予定データ]", Query )

' ************************************************
' 商品集計
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	商品コード NVARCHAR2(4)" & _
"	,経過月 NUMBER(2,0)" & _
"	,当月売上数量 NUMBER(10,0)" & _
"	,当月売上金額 NUMBER(10,0)" & _
"	,更新日 DATE" & _
"	,組織コード NVARCHAR2(4)" & _
"	,primary key(商品コード,経過月)" & _
")"
Call OracleTransfer( "商品集計","[商品集計]", Query )

' ************************************************
' 得意先集計
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	請求先 NVARCHAR2(4)" & _
"	,経過月 NUMBER(2,0)" & _
"	,当月売上金額 NUMBER(10,0)" & _
"	,更新日 DATE" & _
"	,組織コード NVARCHAR2(4)" & _
"	,primary key(請求先,経過月)" & _
")"
Call OracleTransfer( "得意先集計","[得意先集計]", Query )

' ************************************************
' 社員変更履歴
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	社員コード NVARCHAR2(4)" & _
"	,氏名 NVARCHAR2(50)" & _
"	,フリガナ NVARCHAR2(50)" & _
"	,所属 NVARCHAR2(4)" & _
"	,性別 NUMBER(1,0)" & _
"	,作成日 DATE" & _
"	,更新日 DATE" & _
"	,給与 NUMBER" & _
"	,手当 NUMBER" & _
"	,管理者 NVARCHAR2(4)" & _
"	,生年月日 DATE" & _
")"
Call OracleTransfer( "社員変更履歴","[社員変更履歴]", Query )

' ************************************************
' ビュー
' ************************************************
Query = _
"create or replace view V_商品一覧 as" & _
"	SELECT 商品マスタ.商品コード" & _
"	, 商品マスタ.商品名" & _
"	, 商品マスタ.販売単価" & _
"	, 商品分類マスタ.商品分類" & _
"	, 商品分類マスタ.名称 AS 分類名" & _
"	, 商品マスタ.商品区分" & _
"	, コード名称マスタ.名称 AS 区分名" & _
" from" & _
"	(商品マスタ LEFT JOIN 商品分類マスタ" & _
"	ON 商品マスタ.商品分類 = 商品分類マスタ.商品分類" & _
"	) LEFT JOIN コード名称マスタ" & _
"	ON 商品マスタ.商品区分 = コード名称マスタ.コード" & _
" where" & _
"	コード名称マスタ.区分 = 3 and 削除フラグ is NULL"
RunOracle( Query )

Query = _
"create or replace view V_売上日付 as" & _
"	SELECT コントロールマスタ.売上日付" & _
"	FROM コントロールマスタ" & _
"	WHERE コントロールマスタ.キー = '1'"
RunOracle( Query )

Query = _
"create or replace view V_得意先台帳 as" & _
"	SELECT 取引データ.取引先コード" & _
"	, 得意先マスタ.得意先名" & _
"	, 取引データ.取引日付" & _
"	, 取引データ.取引区分" & _
"	, 取引データ.伝票番号" & _
"	, 取引データ.行" & _
"	, 取引データ.商品コード" & _
"	, 商品マスタ.商品名" & _
"	, 取引データ.数量" & _
"	, 取引データ.単価" & _
"	, 取引データ.金額" & _
" from" & _
"	(取引データ INNER JOIN 商品マスタ" & _
"	ON 取引データ.商品コード=商品マスタ.商品コード" & _
"	) INNER JOIN 得意先マスタ" & _
"	ON 取引データ.取引先コード=得意先マスタ.得意先コード" & _
" where" & _
"	取引データ.取引区分 = '10'"
RunOracle( Query )

Query = _
"create or replace view V_社員一覧 as" & _
" select 社員コード" & _
"	,氏名" & _
"	,フリガナ" & _
"	,名称1.名称 as 性別" & _
"	,所属" & _
"	,名称2.名称 as 所属名" & _
" from 社員マスタ" & _
"	,コード名称マスタ 名称1" & _
"	,コード名称マスタ 名称2" & _
" where to_char(性別) = 名称1.コード" & _
"   and 名称1.区分 = 1" & _
"   and 所属 = 名称2.コード" & _
"   and 名称2.区分 = 2"
RunOracle( Query )

Query = _
"create or replace view" & _
"	PROC_ERROR" & _
" as" & _
" select * " & _
" from USER_ERRORS"
RunOracle( Query )

Query = _
"create or replace view" & _
"	PROC_LIST" & _
" as" & _
" select OBJECT_NAME as ""プロシージャ名"" " & _
"	,STATUS as ""状態""  " & _
"	,OBJECT_TYPE  as ""タイプ"" " & _
"	,CREATED  as ""作成日"" " & _
"	,LAST_DDL_TIME  as ""更新日"" " & _
" from USER_OBJECTS " & _
" where OBJECT_TYPE in ('FUNCTION','PROCEDURE') "
RunOracle( Query )

Query = _
"create or replace view" & _
"	PROC_TEXT" & _
" as" & _
" select * from USER_SOURCE"
RunOracle( Query )

' ************************************************
' 終了
' ************************************************

Cn2.Close
Cn.Close

Wscript.Echo "処理が終了しました"

' ************************************************
' Oracle 転送
' ************************************************
function OracleTransfer( strTarget, strTable, QueryCreate )

	Dim Query

	Query = "drop table " & strTarget
	RunOracle( Query )

	RunOracle( QueryCreate )

	Query = "insert into " & strConnectOracle & "." & strDummy & _
	" select * from " & strTable

	RunMdb( Query )

	Query = "alter table " & strDummy & " rename to " & strTarget
	RunOracle( Query )

end function

' ************************************************
' MDB 実行
' ************************************************
function RunMdb( Query )

	on error resume next
	Cn.Execute Query
	if Err.Number <> 0then
		Wscript.Echo Err.Description & vbCrLf & Query
	end if
	on error goto 0

end function

' ************************************************
' Oracle 実行
' ************************************************
function RunOracle( Query )

	on error resume next
	Cn2.Execute Query
	if Err.Number <> 0then
		Wscript.Echo Err.Description & vbCrLf & Query
	end if
	on error goto 0

end function
