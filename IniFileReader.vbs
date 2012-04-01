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

Include "Const.vbs"
Include "Util.vbs"

Class IniFileReader
	Private objUtil
	Private m_strIniFilePath

	'********************************************************************************
	'【　機　能　】：コンストラクタ
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	Private Sub Class_Initialize
		Set objUtil = New Util
	End Sub

	'********************************************************************************
	'【　機　能　】：デストラクタ
	'【　引　数　】：なし
	'【　戻り値　】：なし
	'********************************************************************************
	Private Sub Class_Terminate
		Set objUtil = Nothing
	End Sub

	'********************************************************************************
	'プロパティ定義
	'********************************************************************************
	Public Property Get IniFilePath
		IniFilePath = m_strIniFilePath
	End Property

	Public Property Let IniFilePath(ByVal strIniFilePath)
		On Error Resume Next
		If InStr(strIniFilePath, "\") = 0 Then
			m_strIniFilePath = objUtil.GetScriptDir & "\" & strIniFilePath
		Else
			m_strIniFilePath = strIniFilePath
		End If
	End Property

	'********************************************************************************
	'【　機　能　】：引数の文字列をセクション名の書式にして返す
	'【　引　数　】：strSectionName : セクション名
	'【　戻り値　】：引数の文字列の先頭、末尾にそれぞれ"["、"]"をつけた文字列
	'********************************************************************************
	Private Function GetSection(ByVal strSectionName)
		GetSection = "[" & Trim(strSectionName) & "]"
	End Function

	'********************************************************************************
	'【　機　能　】：引数の文字列から key を返す
	'【　引　数　】：strString : key=value 形式の文字列
	'【　戻り値　】：key となる文字列
	'********************************************************************************
	Private Function GetKey(ByVal strString)
		On Error Resume Next
		GetKey = ""
		If Left(strString, 1) = ";" Then
			Exit Function
		End If
		Dim nIndex
		nIndex = InStr(strString, "=")
		If nIndex = 0 Then
			Exit Function
		End If
		GetKey = objUtil.TrimSpaceTab(Left(strString, nIndex - 1))
	End Function

	'********************************************************************************
	'【　機　能　】：引数の文字列から value を返す
	'【　引　数　】：strString : key=value 形式の文字列
	'【　戻り値　】：value となる文字列
	'********************************************************************************
	Private Function GetValue(ByVal strString)
		On Error Resume Next
		GetValue = ""
		If Left(strString, 1) = ";" Then
			Exit Function
		End If
		Dim nIndex
		nIndex = InStr(strString, "=")
		If nIndex = 0 Then
			Exit Function
		End If
		GetValue = objUtil.TrimSpaceTab(Mid(strString, InStr(strString, "=") + 1))
	End Function

	'********************************************************************************
	'【　機　能　】：iniファイルから引数のセクション、key に該当する value を返す
	'【　引　数　】：strSection : セクション名
	'　　　　　　　：strKey : key
	'              ：strDefaultValue : デフォルト値
	'【　戻り値　】：引数のセクション、key に該当する value。
	'                value が取得できなかった場合は strDefaultValue。
	'********************************************************************************
	Public Function GetIniValue(ByVal strSection, ByVal strKey, ByVal strDefaultValue)
		On Error Resume Next

		GetIniValue = strDefaultValue

		If strSection = "" Or strKey = "" Then
			Exit Function
		End If

		Dim objFso
		Set objFso = CreateObject("Scripting.FileSystemObject")
		Dim objIniFile
		Set objIniFile = objFso.OpenTextFile(iniFilePath, FOR_READING)
		If Err.Number <> 0 Then
			Exit Function
		End If

		Dim strLine
		Dim strValue
		strValue = strDefaultValue
		Do Until objIniFile.AtEndOfStream
			strLine = objIniFile.ReadLine
			If UCase(strLine) = UCase(GetSection(strSection)) Then
				Do Until objIniFile.AtEndOfStream
					strLine = objIniFile.ReadLine
					If Left(strLine, 1) = "[" Then
						Exit Do
					End If
					If UCase(GetKey(strLine)) = UCase(Trim(strKey)) Then
						strValue = GetValue(strLine)
						Exit Do
					End If
				Loop
				Exit Do
			End If
		Loop
		objIniFile.Close
		Set objIniFile = Nothing
		Set objFso = Nothing

		GetIniValue = strValue
	End Function

	'********************************************************************************************
	'【　機　能　】：iniファイルから引数のセクション、
	'　　　　　　　　引数の key と前方一致する key と、該当する value の2次元配列を返す
	'【　引　数　】：strSection : セクション名
	'　　　　　　　：strKey : key
	'【　戻り値　】：引数のセクション、key と前方一致する key と、該当する value の2次元配列。
	'                value が取得できなかった場合は ""。ただし、判定はIsArrayで行ってください。
	'　　　　　　　　GetIniValueByKeyStartsWith(1, n)
	'　　　　　　　　　GetIniValueByKeyStartsWith(0, n)：n個目のキー
	'　　　　　　　　　GetIniValueByKeyStartsWith(1, n)：n個目の値
	'********************************************************************************************
	Public Function GetIniValueByKeyStartsWith(ByVal strSection, ByVal strKey)
		On Error Resume Next
		GetIniValueByKeyStartsWith = GetIniValueByKeyStartsOrEndsWith(strSection, strKey, True)
	End Function

	'********************************************************************************************
	'【　機　能　】：iniファイルから引数のセクション、
	'　　　　　　　　引数の key と後方一致する key と、該当する value の2次元配列を返す
	'【　引　数　】：strSection : セクション名
	'　　　　　　　：strKey : key
	'【　戻り値　】：引数のセクション、key と後方一致する key と、該当する value の2次元配列。
	'                value が取得できなかった場合は ""。ただし、判定はIsArrayで行ってください。
	'　　　　　　　　GetIniValueByKeyEndsWith(1, n)
	'　　　　　　　　　GetIniValueByKeyEndsWith(0, n)：n個目のキー
	'　　　　　　　　　GetIniValueByKeyEndsWith(1, n)：n個目の値
	'********************************************************************************************
	Public Function GetIniValueByKeyEndsWith(ByVal strSection, ByVal strKey)
		On Error Resume Next
		GetIniValueByKeyEndsWith = GetIniValueByKeyStartsOrEndsWith(strSection, strKey, False)
	End Function

	'********************************************************************************************
	'【　機　能　】：iniファイルから引数のセクション、
	'　　　　　　　　引数の key と前方一致または後方一致する key と、該当する value の2次元配列を返す
	'【　引　数　】：strSection : セクション名
	'　　　　　　　：strKey : key
	'              ：bIsStarts : True:前方一致 False:後方一致
	'【　戻り値　】：引数のセクション、key と前方一致または後方一致する key と、該当する value の2次元配列。
	'                value が取得できなかった場合は ""。ただし、判定はIsArrayで行ってください。
	'　　　　　　　　GetIniValueByKeyEndsWith(1, n)
	'　　　　　　　　　GetIniValueByKeyEndsWith(0, n)：n個目のキー
	'　　　　　　　　　GetIniValueByKeyEndsWith(1, n)：n個目の値
	'********************************************************************************************
	Private Function GetIniValueByKeyStartsOrEndsWith(ByVal strSection, ByVal strKey, ByVal bIsStarts)
		On Error Resume Next

		GetIniValueByKeyStartsOrEndsWith = ""

		If strSection = "" Or strKey = "" Then
			Exit Function
		End If

		Dim objFso
		Set objFso = CreateObject("Scripting.FileSystemObject")
		Dim objIniFile
		Set objIniFile = objFso.OpenTextFile(iniFilePath, FOR_READING)
		If Err.Number <> 0 Then
			Exit Function
		End If

		Dim strLine
		Dim strGetKey
		Dim result()
		ReDim result(1, 0)
		Dim nArrayIndex
		nArrayIndex = 0
		Do Until objIniFile.AtEndOfStream
			strLine = objIniFile.ReadLine
			If UCase(strLine) = UCase(GetSection(strSection)) Then
				Do Until objIniFile.AtEndOfStream
					strLine = objIniFile.ReadLine
					If Left(strLine, 1) = "[" Then
						Exit Do
					End If
					If Not Left(strLine, 1) = ";" Then
						strGetKey = GetKey(strLine)
						If (bIsStarts And UCase(Trim(strKey)) = UCase(Left(strGetKey, Len(Trim(strKey))))) Or (Not bIsStarts And UCase(Trim(strKey)) = UCase(Right(strGetKey, Len(Trim(strKey))))) Then
							If nArrayIndex > 0 Then
								ReDim Preserve result(1, nArrayIndex)
							End If
							result(0, nArrayIndex) = strGetKey
							result(1, nArrayIndex) = GetValue(strLine)
							nArrayIndex = nArrayIndex + 1
						End If
					End If
				Loop
				Exit Do
			End If
		Loop
		objIniFile.Close
		Set objIniFile = Nothing
		Set objFso = Nothing
		
		If nArrayIndex > 0 Then
			GetIniValueByKeyStartsOrEndsWith = result
		End If
	End Function

End Class

'********************************************************************************
'【　機　能　】：引数のファイルをインクルードするプロシージャ
'【　引　数　】：strFilePath : インクルードするVBScriptファイル名
'【　戻り値　】：なし
'********************************************************************************
Sub Include(ByVal strFilePath)
	On Error Resume Next

	Dim objFso, objTextStream

    Set objFso = WScript.CreateObject("Scripting.FileSystemObject")

	'読み取り専用でファイルを開く
	Set objTextStream = objFso.OpenTextFile(strFilePath, 1, False)

	'ファイルの内容を実行
	ExecuteGlobal(objTextStream.ReadAll)

    objTextStream.Close
	Set objTextStream = Nothing
	Set objFso = Nothing
End Sub
