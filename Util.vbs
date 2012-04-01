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

'********************************************************************************
'【　機　能　】：ユーティリティクラス
'********************************************************************************
Class Util

	'********************************************************************************
	'【　機　能　】：コンストラクタ
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	private Sub Class_Initialize
	End Sub

	'********************************************************************************
	'【　機　能　】：デストラクタ
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	private Sub Class_Terminate
	End Sub

	'********************************************************************************
	'【　機　能　】：引数で指定されたフォーマットの日時を取得
	'【　引　数　】：日付オブジェクト
	'                フォーマット文字列
	'【　戻り値　】：引数で指定されたフォーマットの日時
	'********************************************************************************
	Public Function FormatDate(dtmSystemDate, strFmtStr)
		On Error Resume Next
		
		FormatDate = vbNullString

		Dim strSystemDate
		strSystemDate = vbNullString
		
		Select Case strFmtStr
			Case "YYYYMM"
				strSystemDate = Year(dtmSystemDate) & _
								Right("0" & Month(dtmSystemDate),2)
			Case "YYYYMMDD"
				strSystemDate = Year(dtmSystemDate) & _
								Right("0" & Month(dtmSystemDate),2) & _
								Right("0" & Day(dtmSystemDate),2)
			Case "YYYY/MM"
				strSystemDate = Year(dtmSystemDate) & "/" & _
								Right("0" & Month(dtmSystemDate),2)
			Case "YYYY/MM/DD"
				strSystemDate = Year(dtmSystemDate) & "/" & _
								Right("0" & Month(dtmSystemDate),2) & "/" & _
								Right("0" & Day(dtmSystemDate),2)
			Case "HH24:MI:SS"
				strSystemDate = Right("0" & Hour(dtmSystemDate),2) & ":" & _
								Right("0" & Minute(dtmSystemDate),2) & ":" & _
								Right("0" & Second(dtmSystemDate),2)
			Case "YYYYMMDDHH24MISS"
				strSystemDate = Year(dtmSystemDate) & _
								Right("0" & Month(dtmSystemDate),2) & _
								Right("0" & Day(dtmSystemDate),2) & _
								Right("0" & Hour(dtmSystemDate),2) & _
								Right("0" & Minute(dtmSystemDate),2) & _
								Right("0" & Second(dtmSystemDate),2)
			Case "HH24:MI"
				strSystemDate = Right("0" & Hour(dtmSystemDate),2) & ":" & _
								Right("0" & Minute(dtmSystemDate),2)
			Case Else
				Err.Raise(5)
		End Select

		FormatDate = strSystemDate
	End Function

	'********************************************************************************
	'【　機　能　】：指定した文字列が、Null値またはEmpty値または
	'                空白(""や半角・全角スペースのみ)の場合に、引数で指定された文字列を返す
	'【　引　数　】：判定する文字列
	'                デフォルト文字列
	'【　戻り値　】：引数がNullまたはEmptyまたは空白の場合は指定された文字列。
	'                それ以外はそのまま
	'********************************************************************************
	Public Function GetDef(ByVal strString, ByVal strDefault)
		On Error Resume Next

		Dim strTmpString
		strTmpString = Trim(strString)

		If IsEmpty(strTmpString) Or IsNull(strTmpString) Or (strTmpString = "") Then
			GetDef = strDefault
		Else
			GetDef = strString
		End If
	End Function

	'********************************************************************************
	'【　機　能　】：スクリプトファイルのディレクトリパスを取得
	'【　引　数　】：なし
	'【　戻り値　】：スクリプトファイルのディレクトリパス(末尾に「\」はつかない)
	'********************************************************************************
	Public Function GetScriptDir
		GetScriptDir = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName) - 1)
	End Function

	'********************************************************************************
	'【　機　能　】：SQLのシングルクォートをエスケープする
	'【　引　数　】：文字列
	'【　戻り値　】：エスケープされた文字列
	'********************************************************************************
	Public Function SqlEscape(ByVal strVal)
		SqlEscape = "'" & Replace(strVal, "'", "''") & "'"
	End Function

	'********************************************************************************
	'【　機　能　】：指定された文字コードでテキストファイル出力を行う
	'【　引　数　】：出力ファイルフルパス, 出力内容, 文字コード(ex. "utf-8")
	'【　戻り値　】：なし
	'********************************************************************************
	Public Sub OutputTextFile(ByVal filePath, ByVal text, ByVal charset)
		On Error Resume Next

		Const adTypeBinary = 1
		Const adTypeText = 2
		Const adSaveCreateOverWrite = 2

		Dim objStream
		Set objStream = WScript.CreateObject("ADODB.Stream")
		objStream.Type = adTypeText
		objStream.Charset = charset
		objStream.Open
		objStream.WriteText(text)

		If UCase(charset) = "UTF-8" Then
			'標準APIではUTF-8にもBOMが出力されるため、
			'最初の3byteをスキップしてバイナリでストリームを読み込んで出力する
			objStream.Position = 0
			objStream.Type = adTypeBinary
			objStream.Position = 3
			Dim bin
			bin = objStream.Read
			objStream.Close
			Set objStream = Nothing

			Set objStream = WScript.CreateObject("ADODB.Stream")
			objStream.Type = adTypeBinary
			objStream.Open
			objStream.Write(bin)
		End If

		objStream.SaveToFile filePath, adSaveCreateOverWrite
		objStream.Close
		Set objStream = Nothing
	End Sub

	'********************************************************************************
	'【　機　能　】：引数の文字列から先頭のスペースまたはタブを除去した文字列を返す
	'【　引　数　】：strString : 文字列
	'【　戻り値　】：引数の文字列から先頭のスペースまたはタブを除去した文字列
	'********************************************************************************
	Public Function LTrimSpaceTab(ByVal strString)
		On Error Resume Next
		Dim strValue
		strValue = LTrim(strString)
		Do While Left(strValue, 1) = vbTab
			strValue = LTrim(Mid(strValue, 2))
		Loop
		LTrimSpaceTab = strValue
	End Function

	'********************************************************************************
	'【　機　能　】：引数の文字列から末尾のスペースまたはタブを除去した文字列を返す
	'【　引　数　】：strString : 文字列
	'【　戻り値　】：引数の文字列から末尾のスペースまたはタブを除去した文字列
	'********************************************************************************
	Public Function RTrimSpaceTab(ByVal strString)
		On Error Resume Next
		Dim strValue
		strValue = RTrim(strString)
		Do While Right(strValue, 1) = vbTab
			strValue = RTrim(Left(strValue, Len(strValue) - 1))
		Loop
		RTrimSpaceTab = strValue
	End Function

	'********************************************************************************
	'【　機　能　】：引数の文字列から先頭と末尾のスペースまたはタブを除去した文字列を返す
	'【　引　数　】：strString : 文字列
	'【　戻り値　】：引数の文字列から先頭と末尾のスペースまたはタブを除去した文字列
	'********************************************************************************
	Public Function TrimSpaceTab(ByVal strString)
		On Error Resume Next
		TrimSpaceTab = RTrimSpaceTab(LTrimSpaceTab(strString))
	End Function

End Class
