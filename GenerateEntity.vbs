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
On Error Resume Next

Include "Logger.vbs"
Include "Util.vbs"
Include "IniFileReader.vbs"
Include "OracleConnectionManager.vbs"

Const ForReading = 1

Dim objFso
Set objFso = CreateObject("Scripting.FileSystemObject")

Dim objLogger
Set objLogger = New Logger

Dim objOracleManager
Set objOracleManager = New OracleConnectionManager

Dim objUtil
Set objUtil = New Util

Dim objIniFileReader
Set objIniFileReader = New IniFileReader

Dim objRecordSet

'エラー発生フラグ
Dim hasError
hasError = False

'ワーニング発生フラグ
Dim hasWarn
hasWarn = False

'出力Javaコードの文字コード
Dim strCharset
'出力Jabaコードの改行コード
Dim strNewLine

'Java予約語の配列
Dim arrayJavaReserved
arrayJavaReserved = Split(JAVA_RESERVED, ",")

main

Set objRecordSet = Nothing
Call objOracleManager.CloseSession

If Err.Number <> 0 Then
	Call WriteLogError(Err, "Oracleデータベース切断に失敗しました。")
End If

Set objFso = Nothing
Set objOracleManager = Nothing
Set objUtil = Nothing
Set objIniFileReader = Nothing
Set objLogger = Nothing

If Err.Number <> 0 Or hasError Then
	WScript.Quit 99
End If

If hasWarn Then
	WScript.Quit 89
End If

WScript.Quit 0

'********************************************************************************
'【　機　能　】：メインプロシージャ
'【　引　数　】：なし
'【　戻り値　】：なし
'********************************************************************************
Sub Main
	On Error Resume Next

    'パラメータ取得(テーブル名)
    Dim strSchemaName
    Dim strTableName

	strSchemaName = ""

    If WScript.Arguments.Count <> 1 Then
		WScript.Echo("Usage : GenerateEntity.vbs [スキーマ名.]テーブル名 | [スキーマ名.]*")
		Exit Sub
    Else
    	Dim strArgs
    	strArgs = Split(WScript.Arguments.Item(0), ".")
    	If UBound(strArgs) = 0 Then
    		strTableName = UCase(strArgs(0))
    	Else
	    	strSchemaName = UCase(strArgs(0))
	    	strTableName = UCase(strArgs(1))
	    End If
    End If

	objLogger.SetLogType("FILE")
	objLogger.SetLogFilePath(objUtil.GetScriptDir & "\debug.log")
'	objLogger.SetLogLevel("DEBUG")

    Call WriteLogDebug("処理開始")

	'iniファイルパス設定
	objIniFileReader.IniFilePath = Trim(GetIniFile)

    'Oracleへ接続
    Call objOracleManager.SetIniFile(Trim(GetIniFile))
    Call objOracleManager.CreateSession
    Call objOracleManager.CreateConnection
    
    If Err.Number <> 0 Then
        Call WriteLogError(Err, "Oracle接続に失敗しました。")
        Exit Sub
    End If

	'oo4o エラーリセット
	objOracleManager.ResetOracleError

	'出力先フォルダ作成
    If Not objFso.FolderExists("dist") Then
        Call objFso.CreateFolder("dist")
        If Err.Number <> 0 Then
        	Call WriteLogError(Err, "フォルダ(dist)の作成に失敗しました。")
        	Exit Sub
        End If
    End If

	'iniファイルより生成ファイルの文字コード、改行コード設定を取得
	strCharset = UCase(objIniFileReader.GetIniValue("DistFileInfo", "Charset", "utf-8"))
	strNewLine = objIniFileReader.GetIniValue("DistFileInfo", "NewLine", vbCrLf)
	If UCase(strNewLine) = "CRLF" Then
		strNewLine = vbCrLf
	ElseIf UCase(strNewLine) = "CR" Then
		strNewLine = vbCr
	ElseIf UCase(strNewLine) = "LF" Then
		strNewLine = vbLf
	End If

	'エンティティ生成
	If strTableName <> "*" Then
		GenerateOneEntity strSchemaName, strTableName
	Else
		GenerateAllEntity strSchemaName
    End If
    
    Call WriteLogDebug("処理終了")
End Sub

'********************************************************************************
'【　機　能　】：ログ出力プロシージャ(DEBUGログ)
'【　引　数　】：メッセージ
'【　戻り値　】：なし
'********************************************************************************
Sub WriteLogDebug(ByVal strMessage)
	On Error Resume Next
    objLogger.Debug(strMessage)
End Sub

'********************************************************************************
'【　機　能　】：ログ出力プロシージャ(INFOログ)
'【　引　数　】：メッセージ
'【　戻り値　】：なし
'********************************************************************************
Sub WriteLogInfo(ByVal strMessage)
	On Error Resume Next
    objLogger.Info(strMessage)
End Sub

'********************************************************************************
'【　機　能　】：ログ出力プロシージャ(WARNログ)
'【　引　数　】：メッセージ
'【　戻り値　】：なし
'********************************************************************************
Sub WriteLogWarn(ByVal strMessage)
	On Error Resume Next

	'ワーニング発生フラグを立てる
    hasWarn = True

	WScript.Echo("[WARN] " & strMessage)
    objLogger.warn(strMessage)
End Sub

'********************************************************************************
'【　機　能　】：ログ出力プロシージャ(ERRORログ)
'【　引　数　】：Errオブジェクト
'                エラーメッセージ
'【　戻り値　】：なし
'********************************************************************************
Sub WriteLogError(ByRef Err, ByVal strMessage)
	On Error Resume Next

	'エラー発生フラグを立てる
	hasError = True

	Dim intErrNumber
	intErrNumber = Err.Number
	Dim strErrDescription
	strErrDescription = Err.Description
	
	Err.Clear

	WScript.Echo("[ERROR] " & strMessage & " Err.Number: " & Err.Number & " Err.Description: " & Err.Description)
	objLogger.Error(strMessage & " Err.Number: " & Err.Number & " Err.Description: " & Err.Description)
End Sub

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

'********************************************************************************
'【　機　能　】：iniファイルパスを取得
'【　引　数　】：なし
'【　戻り値　】：iniファイルパス
'********************************************************************************
Private Function GetIniFile
	GetIniFile = objUtil.GetScriptDir & "\script.ini"
End Function

'********************************************************************************
'【　機　能　】：テーブル名をクラス名に変換
'【　引　数　】：テーブル名
'【　戻り値　】：クラス名
'********************************************************************************
Private Function TableNameToClassName(ByVal strTableName)
	Dim strArray
	Dim strTmpString
	Dim s
	strTmpString = ""
	strArray = Split(LCase(strTableName), "_")
	For Each s In strArray
		strTmpString = strTmpString & Ucase(Left(s, 1)) & Right(s, Len(s) - 1)
	Next
	TableNameToClassName = strTmpString
End Function

'********************************************************************************
'【　機　能　】：カラム名をフィールド名に変換
'【　引　数　】：カラム名
'【　戻り値　】：フィールド名
'********************************************************************************
Private Function ColumnNameToFieldName(ByVal strColumnName)
	Dim result
	result = TableNameToClassName(strColumnName)
	result = LCase(Left(result, 1)) & Right(result, Len(result) - 1)

	'カラム名がJava予約語だった場合は接頭語をつける
	If IsJavaReserved(strColumnName) Then
		result = FIELD_PREFIX_ON_JAVA_RESERVED & UCase(Left(result, 1)) & Right(result, Len(result) - 1)
	End If

	ColumnNameToFieldName = result
End Function

'********************************************************************************
'【　機　能　】：引数の文字列がJava予約語かどうかを返す
'【　引　数　】：文字列
'【　戻り値　】：True:Java予約語、False:Java予約語でない
'********************************************************************************
Private Function IsJavaReserved(ByVal strString)
	Dim result
	result = False

	Dim strJavaReserved
	For Each strJavaReserved In arrayJavaReserved
		If UCase(strString) = UCase(strJavaReserved) Then
			result = True
			Exit For
		End If
	Next

	IsJavaReserved = result
End Function

'********************************************************************************
'【　機　能　】：DBのデータ型をJavaのデータ型に変換
'【　引　数　】：DBのデータ型, スキーマ名, テーブル名, カラム名
'【　戻り値　】：Javaのデータ型
'********************************************************************************
Private Function DBTypeToJavaType(ByVal strDBType, ByVal strSchemaName, ByVal strTableName, ByVal strColumnName)
	Dim result
	result = ""
	If strDBType = "VARCHAR2" Or strDBType = "NVARCHAR2" Or strDBType = "CHAR" Or strDBType = "NCHAR" Or strDBType = "CLOB" Or strDBType = "NCLOB" Or strDBType = "LONG" Then
		result = "String"
	ElseIf strDBType = "NUMBER" Then
		result = DBTypeOfNumberToJavaType(strDBType, strSchemaName, strTableName, strColumnName)
	ElseIf strDBType = "DATE" Or strDBType = "TIMESTAMP" Then
		result = "Date"
	ElseIf strDBType = "BLOB" Then
		result = "byte[]"
	End If
	
	If result = "" Then
		Call WriteLogError(Err, "DBの型(" & strDBType & ")からの変換は未実装です。")
	End If
	
	DBTypeToJavaType = result
End Function

'********************************************************************************
'【　機　能　】：DBの数値データ型をJavaのデータ型に変換
'【　引　数　】：DBの数値データ型, スキーマ名, テーブル名, カラム名
'【　戻り値　】：Javaのデータ型
'********************************************************************************
Private Function DBTypeOfNumberToJavaType(ByVal strDBType, ByVal strSchemaName, ByVal strTableName, ByVal strColumnName)
	Dim result
	result = "Integer"
	
	'該当テーブル.カラムのマッピング先のJava型が、ini ファイルに指定されていないか確認
	If strSchemaName = "" Then
		strSchemaName = objIniFileReader.GetIniValue("DataBaseInfo", "Oracle_User", "")
	End If
	Dim strMappingJavaType
	strMappingJavaType = objIniFileReader.GetIniValue("MappingInfo", strSchemaName & "." & strTableName & "." & strColumnName, "")
	If strMappingJavaType = "" Then
		strMappingJavaType = objIniFileReader.GetIniValue("MappingInfo", strTableName & "." & strColumnName, "")
	End If
	If strMappingJavaType <> "" Then
		If ValidateJavaType(strMappingJavaType) Then
			result = strMappingJavaType
		Else
			Call WriteLogError(Err, "script.iniで指定されたマッピング先のJava型(" & strMappingJavaType & ")には対応していません。")
		End If
	End If
	
	DBTypeOfNumberToJavaType = result
End Function

'********************************************************************************
'【　機　能　】：Javaのデータ型名が正しいか検証する
'【　引　数　】：Javaのデータ型
'【　戻り値　】：True:正しい、False:正しくない
'********************************************************************************
Private Function ValidateJavaType(ByVal strJavaType)
	Dim result
	result = False
	
	Dim javaType
	javaType = Replace(strJavaType, "[]", "")
	
	If javaType = "boolean" Or javaType = "Boolean" Or javaType = "char" Or javaType = "Character" Or javaType = "byte" Or javaType = "Byte" Or javaType = "short" Or javaType = "Short" Or javaType = "int" Or javaType = "Integer" Or javaType = "long" Or javaType = "Long" Or javaType = "float" Or javaType = "Float" Or javaType = "double" Or javaType = "Double" Or javaType = "BigInteger" Or javaType = "BigDecimal" Or javaType = "AtomicBoolean" Or javaType = "AtomicInteger" Or javaType = "AtomicLong" Or javaType = "String" Then
		result = True
	End If
	
	ValidateJavaType = result
End Function

'********************************************************************************
'【　機　能　】：引数のテーブルの主キー自動生成のためのシーケンス名を取得する
'【　引　数　】：スキーマ名, テーブル名
'【　戻り値　】：シーケンス名
'********************************************************************************
Private Function GetSequenceName(ByVal strSchemaName, ByVal strTableName)
	Dim result
	result = ""
	
	'該当テーブルの主キー自動生成のためのシーケンス名が、ini ファイルに指定されていないか確認
	If strSchemaName = "" Then
		strSchemaName = objIniFileReader.GetIniValue("DataBaseInfo", "Oracle_User", "")
	End If
	Dim strSequenceName
	strSequenceName = objIniFileReader.GetIniValue("SequenceInfo", strSchemaName & "." & strTableName, "")
	If strSequenceName = "" Then
		strSequenceName = objIniFileReader.GetIniValue("SequenceInfo", strTableName, "")
	End If
	If strSequenceName <> "" Then
		result = strSequenceName
		If Not ExistsSequence(strSequenceName) Then
			Call WriteLogWarn("script.iniで指定されたシーケンス(" & strSequenceName & ")は存在しません。")
		End If
	End If
	
	GetSequenceName = result
End Function

'********************************************************************************
'【　機　能　】：引数のOracle シーケンスが存在するか検証する
'【　引　数　】：Oracle シーケンス名
'【　戻り値　】：True:存在する、False:存在しない
'********************************************************************************
Private Function ExistsSequence(ByVal strSequenceName)
	On Error Resume Next
	Dim result
	result = False
	
	Dim strSql
	strSql = "select count(*)"
	strSql = strSql & " from USER_SEQUENCES"
	strSql = strSql & " where SEQUENCE_NAME = " & objUtil.SqlEscape(strSequenceName)
	
	Call WriteLogDebug(strSql)
	Call objOracleManager.ExecuteQuery(strSql, objRecordSet)
	
	If Err.Number <> 0 Or objOracleManager.GetOracleErrorCode <> 0 Then
		Call WriteLogError(Err, "SQL実行時エラー（" & strSql & "）。 oo4o: " & objOracleManager.GetOracleErrorCode & " " & objOracleManager.GetOracleErrorMessage)
	Else
		Dim intCount
		intCount = objUtil.GetDef(objRecordSet.Fields(0).Value, "")
		If intCount = 1 Then
			result = True
		End If
	End If
	
	Set objRecordSet = Nothing
	ExistsSequence = result
End Function

'********************************************************************************
'【　機　能　】：引数のテーブル、カラムの関連定義出力文字列を取得する
'【　引　数　】：スキーマ名, テーブル名, カラム名
'【　戻り値　】：関連定義出力文字列
'********************************************************************************
Private Function GetRelationshipText(ByVal strSchemaName, ByVal strTableName, ByVal strColumnName)
	On Error Resume Next
	Dim result
	result = ""

	If strSchemaName = "" Then
		strSchemaName = objIniFileReader.GetIniValue("DataBaseInfo", "Oracle_User", "")
	End If

	Dim bSchemaName
	bSchemaName = True

	Dim strKey
	Dim strValue
	Dim i

	'該当カラムの関連定義(関連の所有者側の定義)が、ini ファイルに指定されていないか確認
	Dim arrayRelationshipInfo
	arrayRelationshipInfo = objIniFileReader.GetIniValueByKeyStartsWith("RelationshipInfo", strSchemaName & "." & strTableName & "." & strColumnName & ":")
	If Not IsArray(arrayRelationshipInfo) Then
		arrayRelationshipInfo = objIniFileReader.GetIniValueByKeyStartsWith("RelationshipInfo", strTableName & "." & strColumnName & ":")
		bSchemaName = False
	End If

	Dim strRelationshipOwnedTableName
	Dim strJoinColumnName
	Dim strReferencedColumnName
	If IsArray(arrayRelationshipInfo) Then
		For i = 0 To UBound(arrayRelationshipInfo, 2)
			strKey = arrayRelationshipInfo(0, i)
			strValue = arrayRelationshipInfo(1, i)
			
			If bSchemaName Then
				strRelationshipOwnedTableName = Split(strKey, ".")(3)
				strJoinColumnName = Split(Split(strKey, ".")(2), ":")(0)
				strReferencedColumnName = Split(strKey, ".")(4)
			Else
				strRelationshipOwnedTableName = Split(Split(strKey, ".")(1), ":")(1)
				strJoinColumnName = Split(Split(strKey, ".")(1), ":")(0)
				strReferencedColumnName = Split(strKey, ".")(2)
			End If
			
			If strValue = "1:1" Then
				result = result & "@OneToOne" & strNewLine
				result = result & "@JoinColumn(name = """ & strJoinColumnName & """, referencedColumnName = """ & strReferencedColumnName & """)" & strNewLine
				result = result & "public " & TableNameToClassName(strRelationshipOwnedTableName) & " " & ColumnNameToFieldName(strRelationshipOwnedTableName) & ";" & strNewLine & strNewLine
			ElseIf strValue = "*:1" Then
				result = result & "@ManyToOne" & strNewLine
				result = result & "@JoinColumn(name = """ & strJoinColumnName & """, referencedColumnName = """ & strReferencedColumnName & """)" & strNewLine
				result = result & "public " & TableNameToClassName(strRelationshipOwnedTableName) & " " & ColumnNameToFieldName(strRelationshipOwnedTableName) & ";" & strNewLine & strNewLine
			Else
				Call WriteLogWarn("script.iniで指定された関連定義(" & strKey & ")の形式が誤っています。1:1、*:1のいずれかで指定してください。")
			End If
		Next
	End If

	'該当カラムの関連定義(関連の被所有者側の定義)が、ini ファイルに指定されていないか確認
	bSchemaName = True
	arrayRelationshipInfo = objIniFileReader.GetIniValueByKeyEndsWith("RelationshipInfo", ":" & strSchemaName & "." & strTableName & "." & strColumnName)
	If Not IsArray(arrayRelationshipInfo) Then
		arrayRelationshipInfo = objIniFileReader.GetIniValueByKeyEndsWith("RelationshipInfo", ":" & strTableName & "." & strColumnName)
		bSchemaName = False
	End If
	
	Dim strRelationshipOwnerTableName
	If IsArray(arrayRelationshipInfo) Then
		For i = 0 To UBound(arrayRelationshipInfo, 2)
			strKey = arrayRelationshipInfo(0, i)
			strValue = arrayRelationshipInfo(1, i)
			
			If bSchemaName Then
				strRelationshipOwnerTableName = Split(strKey, ".")(1)
			Else
				strRelationshipOwnerTableName = Split(strKey, ".")(0)
			End If
			If strValue = "1:1" Then
				result = result & "@OneToOne(mappedBy = """ & ColumnNameToFieldName(strTableName) & """)" & strNewLine
				result = result & "public " & TableNameToClassName(strRelationshipOwnerTableName) & " " & ColumnNameToFieldName(strRelationshipOwnerTableName) & ";" & strNewLine & strNewLine
			ElseIf strValue = "*:1" Then
				result = result & "@OneToMany(mappedBy = """ & ColumnNameToFieldName(strTableName) & """)" & strNewLine
				result = result & "public List<" & TableNameToClassName(strRelationshipOwnerTableName) & "> " & ColumnNameToFieldName(strRelationshipOwnerTableName) & "List;" & strNewLine & strNewLine
			Else
				Call WriteLogWarn("script.iniで指定された関連定義(" & strKey & ")の形式が誤っています。1:1、*:1のいずれかで指定してください。")
			End If
		Next
	End If

	GetRelationshipText = result
End Function

'********************************************************************************
'【　機　能　】：引数のテーブル、カラムの任意の
'                フィールドアノテーション出力文字列を取得する
'【　引　数　】：スキーマ名, テーブル名, カラム名
'【　戻り値　】：フィールドアノテーション出力文字列
'********************************************************************************
Private Function GetFieldAnnotations(ByVal strSchemaName, ByVal strTableName, ByVal strColumnName)
	On Error Resume Next
	Dim result()

	If strSchemaName = "" Then
		strSchemaName = objIniFileReader.GetIniValue("DataBaseInfo", "Oracle_User", "")
	End If

	'該当カラムの任意のフィールドアノテーションが、ini ファイルに指定されていないか確認
	Dim strValue
	strValue = objIniFileReader.GetIniValue("FieldAnnotationInfo", strSchemaName & "." & strTableName & "." & strColumnName, "")
	If strValue = "" Then
		strValue = objIniFileReader.GetIniValue("FieldAnnotationInfo", strTableName & "." & strColumnName, "")
	End If
	If strValue = "" Then
		strValue = objIniFileReader.GetIniValue("FieldAnnotationInfo", strColumnName, "")
	End If
	If strValue <> "" Then
		Dim arrayFieldAnnotations
		arrayFieldAnnotations = Split(strValue, "|")
		Dim f
		Dim nDotIndex
		Dim nArgsIndex
		Dim j
		j = 0
		Dim bImport
		Dim strImport
		Dim strAnnotation

		For Each f In arrayFieldAnnotations
			nArgsIndex = InStr(f, "(")
			If nArgsIndex = 0 Then
				nDotIndex = InStrRev(f, ".")
				If nDotIndex = 0 Then
					'パッケージ指定がない(アノテーションが同一パッケージ)の場合
					strImport = ""
					strAnnotation = "@" & f & strNewLine
				Else
					strImport = "import " & f & ";" & strNewLine
					strAnnotation = "@" & Mid(f, nDotIndex + 1) & strNewLine
				End If
			Else
				bImport = False
				nDotIndex = Len(f)
				Do While InStrRev(f, ".", nDotIndex) > 0
					nDotIndex = InStrRev(f, ".", nDotIndex)
					If nDotIndex > nArgsIndex Then
						nDotIndex = nDotIndex - 1
					Else
						bImport = True
						Exit Do
					End If
				Loop
				If Not bImport Then
					'パッケージ指定がない(アノテーションが同一パッケージ)の場合
					strImport = ""
					strAnnotation = "@" & f & strNewLine
				Else
					strImport = "import " & Left(f, nArgsIndex - 1) & ";" & strNewLine
					strAnnotation = "@" & Mid(f, nDotIndex + 1) & strNewLine
				End If
			End If
			ReDim Preserve result(1, j)
			result(0, j) = strImport
			result(1, j) = strAnnotation
			j = j + 1
		Next
	End If

	GetFieldAnnotations = result
End Function

'********************************************************************************
'【　機　能　】：エンティティを1つ生成する
'【　引　数　】：スキーマ名, テーブル名
'【　戻り値　】：なし
'********************************************************************************
Private Sub GenerateOneEntity(ByVal schemaName, ByVal tableName)
	On Error Resume Next
	Dim strSql
	Dim objRecordSet

	' テーブルコメント取得
	strSql = "select COMMENTS"
	If schemaName = "" Then
		strSql = strSql & " from USER_TAB_COMMENTS"
	Else
		strSql = strSql & " from ALL_TAB_COMMENTS"
	End If
	strSql = strSql & " where TABLE_TYPE = 'TABLE'"
	strSql = strSql & " and TABLE_NAME = " & objUtil.SqlEscape(tableName)
	If schemaName <> "" Then
		strSql = strSql & " and OWNER = " & objUtil.SqlEscape(schemaName)
	End If

	Call WriteLogDebug(strSql)
	Call objOracleManager.ExecuteQuery(strSql, objRecordSet)
    
    If Err.Number <> 0 Or objOracleManager.GetOracleErrorCode <> 0 Then
        Call WriteLogError(Err, "SQL実行時エラー（" & strSql & "）。 oo4o: " & objOracleManager.GetOracleErrorCode & " " & objOracleManager.GetOracleErrorMessage)
		Exit Sub
    End If
    
    If objRecordSet.EOF Then
    	If schemaName <> "" Then
    		WScript.Echo("テーブル " & schemaName & "." & tableName & " は存在しないか権限がありません。")
    	Else
    		WScript.Echo("テーブル " & tableName & " は存在しません。")
    	End If
		Exit Sub
    End If
    
    Dim strTableComment
    strTableComment = objUtil.GetDef(objRecordSet.Fields(0).Value, "")
    Set objRecordSet = Nothing
    
    ' カラム情報取得
	strSql = "select col.COLUMN_NAME, col.DATA_TYPE, com.COMMENTS"
	If schemaName = "" Then
		strSql = strSql & " from USER_TAB_COLUMNS col, USER_COL_COMMENTS com"
	Else
		strSql = strSql & " from ALL_TAB_COLUMNS col, ALL_COL_COMMENTS com"
	End If
	strSql = strSql & " where col.TABLE_NAME = " & objUtil.SqlEscape(tableName)
	strSql = strSql & " and col.TABLE_NAME = com.TABLE_NAME"
	strSql = strSql & " and col.COLUMN_NAME = com.COLUMN_NAME"
	If schemaName <> "" Then
		strSql = strSql & " and col.OWNER = " & objUtil.SqlEscape(schemaName)
		strSql = strSql & " and col.OWNER = com.OWNER"
	End If
	strSql = strSql & " order by col.COLUMN_ID"

	Call WriteLogDebug(strSql)
	Call objOracleManager.ExecuteQuery(strSql, objRecordSet)
    
    If Err.Number <> 0 Or objOracleManager.GetOracleErrorCode <> 0 Then
        Call WriteLogError(Err, "SQL実行時エラー（" & strSql & "）。 oo4o: " & objOracleManager.GetOracleErrorCode & " " & objOracleManager.GetOracleErrorMessage)
		Exit Sub
    End If

	'具象クラスが既に存在するか確認
	'存在する場合は、定義されているフィールドを配列に格納
	Dim strAbstractClassName
    Dim strClassName
    Dim existsClass
    Dim strFields()
	Dim objRE 
	Set objRE = new RegExp
	Dim i
	Dim strLine
   	Dim objClassFile

    ReDim Preserve strFields(0)
    existsClass = False
    strAbstractClassName = "Abstract" & TableNameToClassName(tableName)
    strClassName = TableNameToClassName(tableName)
    If objFso.FileExists("dist\" & strClassName & ".java") Then
    	existsClass = True
	   	Set objClassFile = objFso.OpenTextFile("dist\" & strClassName & ".java", ForReading)
		i = 0
	    While objClassFile.AtEndOfStream = False
			strLine = objClassFile.ReadLine
			objRE.IgnoreCase = False
			objRE.pattern = "^[ \t]*public.*;$"
			If objRE.Test(strLine) Then
				ReDim Preserve strFields(i)
				strFields(i) = strLine
				i = i + 1
			End If
		Wend
		objClassFile.Close
		Set objClassFile = Nothing
	End If

	i = 0
	Dim strColumnName()
	ReDim Preserve strColumnName(0)
	Dim strColumnType()
	ReDim Preserve strColumnType(0)
	Dim strColumnComment()
	ReDim Preserve strColumnComment(0)
	Dim strAllColumnName()
	ReDim Preserve strAllColumnName(0)

	Dim f
	Dim isMatch
	Dim j
	j = 0
    Do While Not objRecordSet.EOF
		ReDim Preserve strAllColumnName(j)
		strAllColumnName(j) = objRecordSet.Fields(0).Value
		j = j + 1
   		isMatch = False
    	If existsClass Then
    		For Each f In strFields
	    		If f <> vbNullString Then
					objRE.IgnoreCase = False
					objRE.pattern = "^[ \t]*public.*[ \t]+" & ColumnNameToFieldName(objRecordSet.Fields(0).Value) & ";$"
					If objRE.Test(f) Then
	    				isMatch = True
	    				Exit For
	    			End If
	    		End If
    		Next
    	End If
    	If Not isMatch Then
	    	ReDim Preserve strColumnName(i)
	    	ReDim Preserve strColumnType(i)
	    	ReDim Preserve strColumnComment(i)
	        strColumnName(i) = objRecordSet.Fields(0).Value
	        strColumnType(i) = objRecordSet.Fields(1).Value
	        strColumnComment(i) = objUtil.GetDef(objRecordSet.Fields(2).Value, "")
	        i = i + 1
	    End If
        Call objRecordSet.DbMoveNext
    Loop

    Set objRecordSet = Nothing

	'テーブルに存在しないフィールドが具象クラスに定義されていないか確認
	Dim strPreLine
	If existsClass Then
		Dim c
		For Each f In strFields
			If f <> vbNullString Then
				isMatch = False
				If Instr(f, "transient") = 0 Then
					For Each c In strAllColumnName
						objRE.IgnoreCase = False
						objRE.pattern = ColumnNameToFieldName(c) & ";$"
						If objRE.Test(f) Then
							isMatch = True
							Exit For
						End If
					Next
					If Not isMatch Then
						'テーブルに存在しないフィールドが永続化対象外として指定されているか確認
						Set objClassFile = objFso.OpenTextFile("dist\" & strClassName & ".java", ForReading)
						strPreLine = ""
						strLine = ""
						Do While objClassFile.AtEndOfStream = False
							strPreLine = strLine
							strLine = objClassFile.ReadLine
							If strLine = f Then
								If Instr(strPreLine, "@Transient") = 0 Then
									WriteLogWarn(strClassName & ".java でテーブルで未定義のフィールドがTransient指定なしに定義されています(" & f & ")")
								End If
								Exit Do
							End If
						Loop
						objClassFile.Close
						Set objClassFile = Nothing
					End If
				End If
			End If
		Next
	End If

    ' プライマリキー情報取得
	strSql = "select col.COLUMN_NAME"
	strSql = strSql & " from USER_CONSTRAINTS cons, USER_CONS_COLUMNS col"
	strSql = strSql & " where cons.TABLE_NAME = col.TABLE_NAME"
	strSql = strSql & " and cons.CONSTRAINT_TYPE = 'P'"
	strSql = strSql & " and cons.CONSTRAINT_NAME = col.CONSTRAINT_NAME"
	strSql = strSql & " and cons.TABLE_NAME = " & objUtil.SqlEscape(tableName)
	strSql = strSql & " and cons.OWNER = col.OWNER"
	If schemaName <> "" Then
		strSql = strSql & " and cons.OWNER = " & objUtil.SqlEscape(schemaName)
	End If

	Call WriteLogDebug(strSql)
	Call objOracleManager.ExecuteQuery(strSql, objRecordSet)
    
    If Err.Number <> 0 Or objOracleManager.GetOracleErrorCode <> 0 Then
        Call WriteLogError(Err, "SQL実行時エラー（" & strSql & "）。 oo4o: " & objOracleManager.GetOracleErrorCode & " " & objOracleManager.GetOracleErrorMessage)
		Exit Sub
    End If

	Dim strPrimaryColumnName()
	ReDim Preserve strPrimaryColumnName(0)

	i = 0
    Do While Not objRecordSet.EOF
   		isMatch = False
    	If existsClass Then
    		For Each f In strFields
    			If f <> vbNullString Then
					objRE.IgnoreCase = False
					objRE.pattern = "^[ \t]*public.*[ \t]+" & ColumnNameToFieldName(objRecordSet.Fields(0).Value) & ";$"
					If objRE.Test(f) Then
	    				isMatch = True
	    				Exit For
	    			End If
	    		End If
    		Next
    	End If
    	If Not isMatch Then
	    	ReDim Preserve strPrimaryColumnName(i)
	        strPrimaryColumnName(i) = objRecordSet.Fields(0).Value
	        i = i + 1
	    End If
        Call objRecordSet.DbMoveNext
    Loop

    Set objRecordSet = Nothing
    Set objRE = Nothing

	'エンティティ生成(抽象クラス)
   	Dim objTemplateFile
   	Set objTemplateFile = objFso.OpenTextFile(objUtil.GetScriptDir & "\abstractTemplate.txt", ForReading)
	Dim strTmpString
	Dim strImport
	strImport = ""
	Dim strText
	strText = ""
	Dim strRelationshipText
	strRelationshipText = ""
	Dim strFieldAnnotations
	Dim k
	
    While objTemplateFile.AtEndOfStream = False
		strLine = objTemplateFile.ReadLine
		If Left(strLine, 7) = "package" Then
			strImport = strImport & strLine & strNewLine
			For i = 0 To UBound(strColumnName)
				If DBTypeToJavaType(strColumnType(i), schemaName, tableName, strColumnName(i)) = "BigInteger" Then
					strImport = strImport & strNewLine
					strImport = strImport & "import java.math.BigInteger;" & strNewLine
					Exit For
				End If
			Next
			For i = 0 To UBound(strColumnName)
				If DBTypeToJavaType(strColumnType(i), schemaName, tableName, strColumnName(i)) = "BigDecimal" Then
					strImport = strImport & strNewLine
					strImport = strImport & "import java.math.BigDecimal;" & strNewLine
					Exit For
				End If
			Next
			For i = 0 To UBound(strColumnName)
				If DBTypeToJavaType(strColumnType(i), schemaName, tableName, strColumnName(i)) = "AtomicBoolean" Then
					strImport = strImport & strNewLine
					strImport = strImport & "import java.util.concurrent.atomic.AtomicBoolean;" & strNewLine
					Exit For
				End If
			Next
			For i = 0 To UBound(strColumnName)
				If DBTypeToJavaType(strColumnType(i), schemaName, tableName, strColumnName(i)) = "AtomicInteger" Then
					strImport = strImport & strNewLine
					strImport = strImport & "import java.util.concurrent.atomic.AtomicInteger;" & strNewLine
					Exit For
				End If
			Next
			For i = 0 To UBound(strColumnName)
				If DBTypeToJavaType(strColumnType(i), schemaName, tableName, strColumnName(i)) = "AtomicLong" Then
					strImport = strImport & strNewLine
					strImport = strImport & "import java.util.concurrent.atomic.AtomicLong;" & strNewLine
					Exit For
				End If
			Next
			
			Dim t
			For Each t In strColumnType
				If t = "DATE" Then
					strImport = strImport & strNewLine
					strImport = strImport & "import java.util.Date;" & strNewLine
					strImport = strImport & "import javax.persistence.Temporal;" & strNewLine
					strImport = strImport & "import javax.persistence.TemporalType;" & strNewLine
					Exit For
				End If
			Next
			
			For Each t In strColumnType
				If t = "CLOB" Or t = "NCLOB" Or t = "LONG" Or t = "BLOB" Then
					strImport = strImport & strNewLine
					strImport = strImport & "import javax.persistence.Basic;" & strNewLine
					strImport = strImport & "import javax.persistence.FetchType;" & strNewLine
					strImport = strImport & "import javax.persistence.Lob;" & strNewLine
					Exit For
				End If
			Next
			
			If strPrimaryColumnName(0) <> "" Then
				strImport = strImport & "import javax.persistence.Id;" & strNewLine
				If GetSequenceName(schemaName, tableName) <> "" Then
					strImport = strImport & "import javax.persistence.GeneratedValue;" & strNewLine
					strImport = strImport & "import javax.persistence.GenerationType;" & strNewLine
					strImport = strImport & "import javax.persistence.SequenceGenerator;" & strNewLine
				End If
			End If
			
			' 関連定義import文出力
			For i = 0 To UBound(strColumnName)
				strRelationshipText = strRelationshipText & GetRelationshipText(schemaName, tableName, strColumnName(i))
			Next
			If strRelationshipText <> "" Then
				If InStr(strRelationshipText, "@OneToOne") > 0 Then
					strImport = strImport & "import javax.persistence.OneToOne;" & strNewLine
				End If
				If InStr(strRelationshipText, "@OneToMany") > 0 Then
					strImport = strImport & "import javax.persistence.OneToMany;" & strNewLine
				End If
				If InStr(strRelationshipText, "@ManyToOne") > 0 Then
					strImport = strImport & "import javax.persistence.ManyToOne;" & strNewLine
				End If
				If InStr(strRelationshipText, "@JoinColumn") > 0 Then
					strImport = strImport & "import javax.persistence.JoinColumn;" & strNewLine
				End If
			End If
			
		ElseIf Left(strLine, 9) = " * Entity" And strTableComment <> "" Then
			strTmpString = Split(strTableComment, " ")
			strText = strText & " * " & strTmpString(0) & " Entity の抽象クラス" & strNewLine
		ElseIf strLine = " * @version" Then
			strText = strText & strLine & " " & objUtil.FormatDate(Now, "YYYY/MM/DD") & strNewLine
		ElseIf Left(strLine, 27) = "public abstract class Clazz" Then
			strText = strText & Replace(strLine, "Clazz", strAbstractClassName) & strNewLine
		ElseIf Left(strLine, 1) = "}" Then
			For i = 0 To UBound(strColumnName)
				'フィールドコメント
				If strColumnComment(i) <> "" Then
					strText = strText & vbTab & "/** " & Replace(strColumnComment(i), strNewLine, strNewLine & vbTab) & " */" & strNewLine
				End If
				For j = 0 To UBound(strPrimaryColumnName)
					If strColumnName(i) = strPrimaryColumnName(j) Then
						strText = strText & vbTab & "@Id" & strNewLine
						If GetSequenceName(schemaName, tableName) <> "" Then
							strText = strText & vbTab & "@GeneratedValue(strategy = GenerationType.SEQUENCE, generator = """ & tableName & "_GEN"")" & strNewLine
							strText = strText & vbTab & "@SequenceGenerator(name = """ & tableName & "_GEN"", sequenceName = """ & GetSequenceName(schemaName, tableName) & """, allocationSize = 1)" & strNewLine
						End If
					End If
				Next
				If strColumnType(i) = "DATE" Or strColumnType(i) = "TIMESTAMP" Then
					strText = strText & vbTab & "@Temporal(TemporalType.TIMESTAMP)" & strNewLine
				End If
				If strColumnType(i) = "CLOB" Or strColumnType(i) = "NCLOB" Or strColumnType(i) = "LONG" Or strColumnType(i) = "BLOB" Then
					strText = strText & vbTab & "@Basic(fetch = FetchType.LAZY)" & strNewLine & vbTab & "@Lob" & strNewLine
				End If
				
				'任意のフィールドアノテーション出力
				strFieldAnnotations = GetFieldAnnotations(schemaName, tableName, strColumnName(i))
				If IsArray(strFieldAnnotations) Then
					For k = 0 To UBound(strFieldAnnotations, 2)
						strImport = strImport & strFieldAnnotations(0, k)
						strText = strText & vbTab & strFieldAnnotations(1, k)
					Next
				End If
				
				'カラム名がJava予約語だった場合は @Column アノテーションをつける
				If IsJavaReserved(strColumnName(i)) Then
					strImport = strImport & "import javax.persistence.Column;"
					strText = strText & vbTab & "@Column(name = """ & strColumnName(i) & """)" & strNewLine
				End If
				
				'フィールド出力
				strText = strText & vbTab & "public " & DBTypeToJavaType(strColumnType(i), schemaName, tableName, strColumnName(i)) & " " & ColumnNameToFieldName(strColumnName(i)) & ";" & strNewLine & strNewLine
			Next
			
			' 関連定義出力
			If strRelationshipText <> "" Then
				strText = strText & vbTab & Replace(Replace(strRelationshipText, strNewLine, strNewLine & vbTab), vbTab & strNewLine, strNewLine)
				'末尾の vbTab を削除
				strText = Left(strText, Len(strText) - 1)
			End If
			
			' クラスの終了カッコ "}" 出力
			strText = strText & strLine & strNewLine
		Else
			strText = strText & strLine & strNewLine
		End If
	Wend
	objTemplateFile.Close
	Set objTemplateFile = Nothing

	'ファイル出力
	objUtil.OutputTextFile "dist\" & strAbstractClassName & ".java", strImport & strText, strCharset

	'エンティティ生成(具象クラス)
    If Not existsClass Then
	   	Set objTemplateFile = objFso.OpenTextFile(objUtil.GetScriptDir & "\template.txt", ForReading)
		strText = ""
	    While objTemplateFile.AtEndOfStream = False
			strLine = objTemplateFile.ReadLine
			If Left(strLine, 7) = "package" Then
				strText = strText & strLine & strNewLine
			ElseIf strLine = " * Entity" And strTableComment <> "" Then
				strTmpString = Split(strTableComment, " ")
				strText = strText & " * " & strTmpString(0) & " Entity" & strNewLine
			ElseIf strLine = " * @version" Then
				strText = strText & strLine & " " & objUtil.FormatDate(Now, "YYYY/MM/DD") & strNewLine
			ElseIf strLine = "@Table" And schemaName <> "" Then
				strText = strText & strLine & "(schema = """ & schemaName & """)" & strNewLine
			ElseIf Left(strLine, 18) = "public class Clazz" Then
				strText = strText & "public class " & strClassName & " extends " & strAbstractClassName & " {" & strNewLine
			Else
				strText = strText & strLine & strNewLine
			End If
		Wend
		objTemplateFile.Close
		Set objTemplateFile = Nothing

		'ファイル出力
		objUtil.OutputTextFile "dist\" & strClassName & ".java", strText, strCharset
	End If
End Sub

'********************************************************************************
'【　機　能　】：エンティティをすべて生成する
'【　引　数　】：スキーマ名
'【　戻り値　】：なし
'********************************************************************************
Private Sub GenerateAllEntity(ByVal schemaName)
	On Error Resume Next
	Dim strSql
	Dim objRecordSet

	' テーブルリスト取得
	strSql = "select TABLE_NAME"
	If schemaName = "" Then
		strSql = strSql & " from USER_TABLES"
	Else
		strSql = strSql & " from ALL_TABLES"
	End If
	If schemaName <> "" Then
		strSql = strSql & " where OWNER = " & objUtil.SqlEscape(schemaName)
	End If

	Call WriteLogDebug(strSql)
	Call objOracleManager.ExecuteQuery(strSql, objRecordSet)
    
    If Err.Number <> 0 Or objOracleManager.GetOracleErrorCode <> 0 Then
        Call WriteLogError(Err, "SQL実行時エラー（" & strSql & "）。 oo4o: " & objOracleManager.GetOracleErrorCode & " " & objOracleManager.GetOracleErrorMessage)
		Exit Sub
    End If
    
    If objRecordSet.EOF Then
   		WScript.Echo("対象のテーブルが存在しません。")
		Exit Sub
    End If
    
    Dim tableList()
    Dim i
    i = 0
    Do While Not objRecordSet.EOF
    	ReDim Preserve tableList(i)
        tableList(i) = objRecordSet.Fields(0).Value
        i = i + 1
        Call objRecordSet.DbMoveNext
    Loop
    Set objRecordSet = Nothing

	Dim table
	For Each table In tableList
		GenerateOneEntity schemaName, table
	Next
End Sub
