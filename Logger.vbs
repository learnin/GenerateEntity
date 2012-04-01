'/*
'* Copyright 2012 Manabu Inoue
'*
'* Licensed under the Apache License, Version 2.0 (the "License");
'* you may not use this file except in compliance with the License.
'* You may obtain a copy of the License at
'*
'* http://www.apache.org/licenses/LICENSE-2.0
'*
'* Unless required by applicable law or agreed to in writing, software
'* distributed under the License is distributed on an "AS IS" BASIS,
'* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'* See the License for the specific language governing permissions and
'* limitations under the License.
'*/
Option Explicit

Include "Util.vbs"

'ログレベル定義
Const Logger_LOG_LEVEL_DEBUG = 1
Const Logger_LOG_LEVEL_INFO = 2
Const Logger_LOG_LEVEL_WARN = 3
Const Logger_LOG_LEVEL_ERROR = 4
Const Logger_LOG_LEVEL_FATAL = 5

'ログ出力先定義
Const Logger_LOG_TYPE_FILE = 1
Const Logger_LOG_TYPE_EVENTLOG = 2

'********************************************************************************
'【　機　能　】：ログ管理クラス
'********************************************************************************
Class Logger
	private intLogLevel
	private intLogType
	private objWshShell
	Private objFso
	private strLogFilePath
	private objUtil

	'********************************************************************************
	'【　機　能　】：コンストラクタ
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	private Sub Class_Initialize
		intLogLevel = Logger_LOG_LEVEL_INFO
		intLogType = Logger_LOG_TYPE_FILE

		Set objWshShell = WScript.CreateObject("WScript.Shell")
		Set objFso = CreateObject("Scripting.FileSystemObject")
		Set objUtil = New Util
	End Sub

	'********************************************************************************
	'【　機　能　】：デストラクタ
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	private Sub Class_Terminate
		On Error Resume Next
		Set objWshShell = Nothing
		Set objFso = Nothing
		Set objUtil = Nothing
	End Sub

	'********************************************************************************
	'【　機　能　】：ログレベルをセットする
	'【　引　数　】：ログレベル文字列
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub SetLogLevel(ByVal strLogLevel)
		If strLogLevel = "DEBUG" Then
			intLogLevel = Logger_LOG_LEVEL_DEBUG
		ElseIf strLogLevel = "INFO" Then
			intLogLevel = Logger_LOG_LEVEL_INFO
		ElseIf strLogLevel = "WARN" Then
			intLogLevel = Logger_LOG_LEVEL_WARN
		ElseIf strLogLevel = "ERROR" Then
			intLogLevel = Logger_LOG_LEVEL_ERROR
		ElseIf strLogLevel = "FATAL" Then
			intLogLevel = Logger_LOG_LEVEL_FATAL
		Else
			Err.Raise(5)
		End If
	End Sub

	'********************************************************************************
	'【　機　能　】：ログ種別をセットする
	'【　引　数　】：ログ種別
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub SetLogType(ByVal strLogType)
		If strLogType = "FILE" Then
			intLogType = Logger_LOG_TYPE_FILE
		ElseIf strLogType = "EVENTLOG" Then
			intLogType = Logger_LOG_TYPE_EVENTLOG
		Else
			Err.Raise(5)
		End If
	End Sub

	'********************************************************************************
	'【　機　能　】：ログファイルパスをセットする
	'【　引　数　】：ログファイルフルパス
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub SetLogFilePath(ByVal logFilePath)
		strLogFilePath = logFilePath
	End Sub

	'********************************************************************************
	'【　機　能　】：DEBUGレベルが有効か判定する
	'【　引　数　】：なし
	'【　戻り値　】：有効ならTrue、無効ならFalse
	'********************************************************************************
	Public Function IsDebugEnabled
		If intLogLevel <= Logger_LOG_LEVEL_DEBUG Then
			IsDebugEnabled = True
		Else
			IsDebugEnabled = False
		End If
	End Function

	'********************************************************************************
	'【　機　能　】：INFOレベルが有効か判定する
	'【　引　数　】：なし
	'【　戻り値　】：有効ならTrue、無効ならFalse
	'********************************************************************************
	Public Function IsInfoEnabled
		If intLogLevel <= Logger_LOG_LEVEL_INFO Then
			IsInfoEnabled = True
		Else
			IsInfoEnabled = False
		End If
	End Function

	'********************************************************************************
	'【　機　能　】：WARNレベルが有効か判定する
	'【　引　数　】：なし
	'【　戻り値　】：有効ならTrue、無効ならFalse
	'********************************************************************************
	Public Function IsWarnEnabled
		If intLogLevel <= Logger_LOG_LEVEL_WARN Then
			IsWarnEnabled = True
		Else
			IsWarnEnabled = False
		End If
	End Function

	'********************************************************************************
	'【　機　能　】：ERRORレベルが有効か判定する
	'【　引　数　】：なし
	'【　戻り値　】：有効ならTrue、無効ならFalse
	'********************************************************************************
	Public Function IsErrorEnabled
		If intLogLevel <= Logger_LOG_LEVEL_ERROR Then
			IsErrorEnabled = True
		Else
			IsErrorEnabled = False
		End If
	End Function

	'********************************************************************************
	'【　機　能　】：FATALレベルが有効か判定する
	'【　引　数　】：なし
	'【　戻り値　】：有効ならTrue、無効ならFalse
	'********************************************************************************
	Public Function IsFatalEnabled
		If intLogLevel <= Logger_LOG_LEVEL_FATAL Then
			IsFatalEnabled = True
		Else
			IsFatalEnabled = False
		End If
	End Function

	'********************************************************************************
	'【　機　能　】：DEBUGログを出力
	'【　引　数　】：エラーメッセージ
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub Debug(ByVal strMessage)
		On Error Resume Next
		If IsDebugEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_DEBUG, strMessage)
		End If
	End Sub

	'********************************************************************************
	'【　機　能　】：INFOログを出力
	'【　引　数　】：エラーメッセージ
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub Info(ByVal strMessage)
		On Error Resume Next
		If IsInfoEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_INFO, strMessage)
		End If
	End Sub

	'********************************************************************************
	'【　機　能　】：WARNログを出力
	'【　引　数　】：エラーメッセージ
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub Warn(ByVal strMessage)
		On Error Resume Next
		If IsWarnEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_WARN, strMessage)
		End If
	End Sub

	'********************************************************************************
	'【　機　能　】：ERRORログを出力
	'【　引　数　】：エラーメッセージ
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub Error(ByVal strMessage)
		On Error Resume Next
		If IsErrorEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_ERROR, strMessage)
		End If
	End Sub

	'********************************************************************************
	'【　機　能　】：FATALログを出力
	'【　引　数　】：エラーメッセージ
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub Fatal(ByVal strMessage)
		On Error Resume Next
		If IsFatalEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_FATAL, strMessage)
		End If
	End Sub

	'********************************************************************************
	'【　機　能　】：ログ種別により指定のログを出力
	'【　引　数　】：ログレベル
	'                エラーメッセージ
	'【　戻り値　】：なし
	'********************************************************************************
	Private Sub WriteLog(ByVal logLevel, ByVal strMessage)
		On Error Resume Next
		If intLogType = Logger_LOG_TYPE_FILE Then
			Call WriteLogFile(logLevel, strMessage)
		ElseIf intLogType = Logger_LOG_TYPE_EVENTLOG Then
			Call WriteEventLog(logLevel, strMessage)
		End If
	End Sub

	'********************************************************************************
	'【　機　能　】：イベントログを出力
	'【　引　数　】：ログレベル
	'                エラーメッセージ
	'【　戻り値　】：なし
	'********************************************************************************
	Private Sub WriteEventLog(ByVal logLevel, ByVal strMessage)
		On Error Resume Next

		Dim intEventLogLevel

		'Loggerのログレベルとイベントログのログレベルのマッピング
		If logLevel = Logger_LOG_LEVEL_DEBUG Then
			'情報
			intEventLogLevel = 4
		ElseIf logLevel = Logger_LOG_LEVEL_INFO Then
			'情報
			intEventLogLevel = 4
		ElseIf logLevel = Logger_LOG_LEVEL_WARN Then
			'警告
			intEventLogLevel = 2
		ElseIf logLevel = Logger_LOG_LEVEL_ERROR Then
			'エラー
			intEventLogLevel = 1
		ElseIf logLevel = Logger_LOG_LEVEL_FATAL Then
			'エラー
			intEventLogLevel = 1
		End If

		objWshShell.LogEvent intEventLogLevel, strMessage
	End Sub

	'********************************************************************************
	'【　機　能　】：ログファイルを出力
	'【　引　数　】：ログレベル
	'                エラーメッセージ
	'【　戻り値　】：なし
	'********************************************************************************
	Private Sub WriteLogFile(ByVal logLevel, ByVal strMessage)
		On Error Resume Next

		Const ForReading = 1, ForAppending = 8
	    Dim dateNow
	    dateNow = Now
		
		Dim objLogFile
		Set objLogFile = objFso.OpenTextFile(strLogFilePath, ForAppending, True)
		Dim strLogLevel
		If logLevel = Logger_LOG_LEVEL_DEBUG Then
			strLogLevel = "DEBUG"
		ElseIf logLevel = Logger_LOG_LEVEL_INFO Then
			strLogLevel = "INFO"
		ElseIf logLevel = Logger_LOG_LEVEL_WARN Then
			strLogLevel = "WARN"
		ElseIf logLevel = Logger_LOG_LEVEL_ERROR Then
			strLogLevel = "ERROR"
		ElseIf logLevel = Logger_LOG_LEVEL_FATAL Then
			strLogLevel = "FATAL"
		End If

		objLogFile.WriteLine(objUtil.FormatDate(dateNow, "YYYY/MM/DD") & Space(1) & objUtil.FormatDate(dateNow, "HH24:MI:SS") & " [" & strLogLevel & "] " & strMessage)
	    objLogFile.Close
		Set objLogFile = Nothing
	End Sub
End Class

'********************************************************************************
'【　機　能　】：引数のファイルをインクルードするプロシージャ
'【　引　数　】：インクルードするVBScriptファイル名
'【　戻り値　】：なし
'********************************************************************************
Sub Include(ByVal FilePath)
	On Error Resume Next

	Dim objFso, objTextStream

    Set objFso = WScript.CreateObject("Scripting.FileSystemObject")

	'読み取り専用でファイルを開く
	Set objTextStream = objFso.OpenTextFile(FilePath, "1", False)

	'ファイルの内容を実行
	ExecuteGlobal(objTextStream.ReadAll)

    objTextStream.Close
	Set objTextStream = Nothing
	Set objFso = Nothing
End Sub
