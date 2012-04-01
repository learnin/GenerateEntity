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

Include "IniFileReader.vbs"

'********************************************************************************
'【　機　能　】：oo4oを使用したOracle接続管理を行うクラス
'********************************************************************************
Class OracleConnectionManager
	Private objSession
	Private objConnection
	Private objIniFileReader

	'********************************************************************************
	'【　機　能　】：コンストラクタ
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	Private Sub Class_Initialize
		Set objSession = Nothing
		Set objConnection = Nothing
		Set objIniFileReader = New IniFileReader
	End Sub

	'********************************************************************************
	'【　機　能　】：デストラクタ
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	Private Sub Class_Terminate
		On Error Resume Next
		Call CloseSession
		Set objIniFileReader = Nothing
	End Sub

	'********************************************************************************
	'【　機　能　】：iniファイルパスをセット
	'【　引　数　】：iniファイルパス
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub SetIniFile(ByVal strIniFilePath)
		objIniFileReader.IniFilePath = strIniFilePath
	End Sub

	'********************************************************************************
	'【　機　能　】：oo4oセッションを取得する
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub CreateSession
		On Error Resume Next
		Set objSession = CreateObject("OracleInProcServer.XOraSession")
		
		If Err.Number <> 0 Then
			Set objSession = Nothing
			Exit Sub
		End If
	End Sub

	'********************************************************************************
	'【　機　能　】：iniファイルより接続情報を取得し、oo4o接続を行う
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub CreateConnection
		On Error Resume Next
		Const ORADB_DEFAULT = &H0
		
    	Set objConnection = objSession.OpenDatabase(objIniFileReader.GetIniValue("DataBaseInfo", "Oracle_Sid", vbNullString), objIniFileReader.GetIniValue("DataBaseInfo", "Oracle_User", vbNullString) & "/" & objIniFileReader.GetIniValue("DataBaseInfo", "Oracle_Passwd", vbNullString), ORADB_DEFAULT)
    	
    	If Err.Number <> 0 Then
    		Set objConnection = Nothing
    		Exit Sub
    	End If
	End Sub

	'********************************************************************************
	'【　機　能　】：oo4oエラーリセットを行う
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub ResetOracleError
		On Error Resume Next
        objConnection.LastServerErrReset
	End Sub

	'********************************************************************************
	'【　機　能　】：oo4oエラーコードを取得する
	'【　引　数　】：なし
	'【　戻り値　】：oo4oエラーコード
	'********************************************************************************
	Public Function GetOracleErrorCode
		On Error Resume Next
        GetOracleErrorCode = objConnection.LastServerErr
	End Function

	'********************************************************************************
	'【　機　能　】：oo4oエラーメッセージを取得する
	'【　引　数　】：なし
	'【　戻り値　】：oo4oエラーメッセージ
	'********************************************************************************
	Public Function GetOracleErrorMessage
		On Error Resume Next
        GetOracleErrorMessage = objConnection.LastServerErrText
	End Function

	'********************************************************************************
	'【　機　能　】：トランザクションを開始する
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub BeginTrans
		On Error Resume Next
        objSession.BeginTrans
	End Sub

	'********************************************************************************
	'【　機　能　】：トランザクションをコミットする
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub CommitTrans
		On Error Resume Next
        objSession.CommitTrans
	End Sub

	'********************************************************************************
	'【　機　能　】：トランザクションをロールバックする
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub Rollback
		On Error Resume Next
        objSession.Rollback
	End Sub

	'********************************************************************************
	'【　機　能　】：更新系SQLを実行する
	'【　引　数　】：SQL
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub ExecuteSQL(ByVal strSql)
		On Error Resume Next
        objConnection.ExecuteSQL(strSql)
	End Sub

	'********************************************************************************
	'【　機　能　】：参照系SQLを実行する
	'【　引　数　】：SQL
	'【　戻り値　】：結果セット
	'********************************************************************************
	Public Sub ExecuteQuery(ByVal strSql, ByRef RecordSet)
		On Error Resume Next
		
		Const ORADYN_DEFAULT       = &H0
		Const ORADYN_NO_AUTOBIND   = &H1
		Const ORADYN_NO_BLANKSTRIP = &H2
		Const ORADYN_READONLY      = &H4
		Const ORADYN_NOCACHE       = &H8
		Const ORADYN_ORAMODE       = &H10
		Const ORADYN_DBDEFAULT     = &H20
		Const ORADYN_NO_MOVEFIRST  = &H40
		Const ORADYN_DIRTY_WRITE   = &H80
		
        Set RecordSet = objConnection.CreateDynaset(strSql, ORADYN_READONLY)
	End Sub

	'********************************************************************************
	'【　機　能　】：oo4oセッション、接続を閉じる
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub CloseSession
		On Error Resume Next
		
        Set objConnection = Nothing
		Set objSession = Nothing
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
