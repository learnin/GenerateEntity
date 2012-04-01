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

'�G���[�����t���O
Dim hasError
hasError = False

'���[�j���O�����t���O
Dim hasWarn
hasWarn = False

'�o��Java�R�[�h�̕����R�[�h
Dim strCharset
'�o��Jaba�R�[�h�̉��s�R�[�h
Dim strNewLine

'Java�\���̔z��
Dim arrayJavaReserved
arrayJavaReserved = Split(JAVA_RESERVED, ",")

main

Set objRecordSet = Nothing
Call objOracleManager.CloseSession

If Err.Number <> 0 Then
	Call WriteLogError(Err, "Oracle�f�[�^�x�[�X�ؒf�Ɏ��s���܂����B")
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
'�y�@�@�@�\�@�z�F���C���v���V�[�W��
'�y�@���@���@�z�F�Ȃ�
'�y�@�߂�l�@�z�F�Ȃ�
'********************************************************************************
Sub Main
	On Error Resume Next

    '�p�����[�^�擾(�e�[�u����)
    Dim strSchemaName
    Dim strTableName

	strSchemaName = ""

    If WScript.Arguments.Count <> 1 Then
		WScript.Echo("Usage : GenerateEntity.vbs [�X�L�[�}��.]�e�[�u���� | [�X�L�[�}��.]*")
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

    Call WriteLogDebug("�����J�n")

	'ini�t�@�C���p�X�ݒ�
	objIniFileReader.IniFilePath = Trim(GetIniFile)

    'Oracle�֐ڑ�
    Call objOracleManager.SetIniFile(Trim(GetIniFile))
    Call objOracleManager.CreateSession
    Call objOracleManager.CreateConnection
    
    If Err.Number <> 0 Then
        Call WriteLogError(Err, "Oracle�ڑ��Ɏ��s���܂����B")
        Exit Sub
    End If

	'oo4o �G���[���Z�b�g
	objOracleManager.ResetOracleError

	'�o�͐�t�H���_�쐬
    If Not objFso.FolderExists("dist") Then
        Call objFso.CreateFolder("dist")
        If Err.Number <> 0 Then
        	Call WriteLogError(Err, "�t�H���_(dist)�̍쐬�Ɏ��s���܂����B")
        	Exit Sub
        End If
    End If

	'ini�t�@�C����萶���t�@�C���̕����R�[�h�A���s�R�[�h�ݒ���擾
	strCharset = UCase(objIniFileReader.GetIniValue("DistFileInfo", "Charset", "utf-8"))
	strNewLine = objIniFileReader.GetIniValue("DistFileInfo", "NewLine", vbCrLf)
	If UCase(strNewLine) = "CRLF" Then
		strNewLine = vbCrLf
	ElseIf UCase(strNewLine) = "CR" Then
		strNewLine = vbCr
	ElseIf UCase(strNewLine) = "LF" Then
		strNewLine = vbLf
	End If

	'�G���e�B�e�B����
	If strTableName <> "*" Then
		GenerateOneEntity strSchemaName, strTableName
	Else
		GenerateAllEntity strSchemaName
    End If
    
    Call WriteLogDebug("�����I��")
End Sub

'********************************************************************************
'�y�@�@�@�\�@�z�F���O�o�̓v���V�[�W��(DEBUG���O)
'�y�@���@���@�z�F���b�Z�[�W
'�y�@�߂�l�@�z�F�Ȃ�
'********************************************************************************
Sub WriteLogDebug(ByVal strMessage)
	On Error Resume Next
    objLogger.Debug(strMessage)
End Sub

'********************************************************************************
'�y�@�@�@�\�@�z�F���O�o�̓v���V�[�W��(INFO���O)
'�y�@���@���@�z�F���b�Z�[�W
'�y�@�߂�l�@�z�F�Ȃ�
'********************************************************************************
Sub WriteLogInfo(ByVal strMessage)
	On Error Resume Next
    objLogger.Info(strMessage)
End Sub

'********************************************************************************
'�y�@�@�@�\�@�z�F���O�o�̓v���V�[�W��(WARN���O)
'�y�@���@���@�z�F���b�Z�[�W
'�y�@�߂�l�@�z�F�Ȃ�
'********************************************************************************
Sub WriteLogWarn(ByVal strMessage)
	On Error Resume Next

	'���[�j���O�����t���O�𗧂Ă�
    hasWarn = True

	WScript.Echo("[WARN] " & strMessage)
    objLogger.warn(strMessage)
End Sub

'********************************************************************************
'�y�@�@�@�\�@�z�F���O�o�̓v���V�[�W��(ERROR���O)
'�y�@���@���@�z�FErr�I�u�W�F�N�g
'                �G���[���b�Z�[�W
'�y�@�߂�l�@�z�F�Ȃ�
'********************************************************************************
Sub WriteLogError(ByRef Err, ByVal strMessage)
	On Error Resume Next

	'�G���[�����t���O�𗧂Ă�
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
'�y�@�@�@�\�@�z�F�����̃t�@�C�����C���N���[�h����v���V�[�W��
'�y�@���@���@�z�F�C���N���[�h����VBScript�t�@�C����
'�y�@�߂�l�@�z�F�Ȃ�
'********************************************************************************
Sub Include(ByVal FilePath)
	On Error Resume Next

	Dim objFso, objTextStream

    Set objFso = WScript.CreateObject("Scripting.FileSystemObject")

	'�ǂݎ���p�Ńt�@�C�����J��
	Set objTextStream = objFso.OpenTextFile(FilePath, "1", False)

	'�t�@�C���̓��e�����s
	ExecuteGlobal(objTextStream.ReadAll)

    objTextStream.Close
	Set objTextStream = Nothing
	Set objFso = Nothing
End Sub

'********************************************************************************
'�y�@�@�@�\�@�z�Fini�t�@�C���p�X���擾
'�y�@���@���@�z�F�Ȃ�
'�y�@�߂�l�@�z�Fini�t�@�C���p�X
'********************************************************************************
Private Function GetIniFile
	GetIniFile = objUtil.GetScriptDir & "\script.ini"
End Function

'********************************************************************************
'�y�@�@�@�\�@�z�F�e�[�u�������N���X���ɕϊ�
'�y�@���@���@�z�F�e�[�u����
'�y�@�߂�l�@�z�F�N���X��
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
'�y�@�@�@�\�@�z�F�J���������t�B�[���h���ɕϊ�
'�y�@���@���@�z�F�J������
'�y�@�߂�l�@�z�F�t�B�[���h��
'********************************************************************************
Private Function ColumnNameToFieldName(ByVal strColumnName)
	Dim result
	result = TableNameToClassName(strColumnName)
	result = LCase(Left(result, 1)) & Right(result, Len(result) - 1)

	'�J��������Java�\��ꂾ�����ꍇ�͐ړ��������
	If IsJavaReserved(strColumnName) Then
		result = FIELD_PREFIX_ON_JAVA_RESERVED & UCase(Left(result, 1)) & Right(result, Len(result) - 1)
	End If

	ColumnNameToFieldName = result
End Function

'********************************************************************************
'�y�@�@�@�\�@�z�F�����̕�����Java�\��ꂩ�ǂ�����Ԃ�
'�y�@���@���@�z�F������
'�y�@�߂�l�@�z�FTrue:Java�\���AFalse:Java�\���łȂ�
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
'�y�@�@�@�\�@�z�FDB�̃f�[�^�^��Java�̃f�[�^�^�ɕϊ�
'�y�@���@���@�z�FDB�̃f�[�^�^, �X�L�[�}��, �e�[�u����, �J������
'�y�@�߂�l�@�z�FJava�̃f�[�^�^
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
		Call WriteLogError(Err, "DB�̌^(" & strDBType & ")����̕ϊ��͖������ł��B")
	End If
	
	DBTypeToJavaType = result
End Function

'********************************************************************************
'�y�@�@�@�\�@�z�FDB�̐��l�f�[�^�^��Java�̃f�[�^�^�ɕϊ�
'�y�@���@���@�z�FDB�̐��l�f�[�^�^, �X�L�[�}��, �e�[�u����, �J������
'�y�@�߂�l�@�z�FJava�̃f�[�^�^
'********************************************************************************
Private Function DBTypeOfNumberToJavaType(ByVal strDBType, ByVal strSchemaName, ByVal strTableName, ByVal strColumnName)
	Dim result
	result = "Integer"
	
	'�Y���e�[�u��.�J�����̃}�b�s���O���Java�^���Aini �t�@�C���Ɏw�肳��Ă��Ȃ����m�F
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
			Call WriteLogError(Err, "script.ini�Ŏw�肳�ꂽ�}�b�s���O���Java�^(" & strMappingJavaType & ")�ɂ͑Ή����Ă��܂���B")
		End If
	End If
	
	DBTypeOfNumberToJavaType = result
End Function

'********************************************************************************
'�y�@�@�@�\�@�z�FJava�̃f�[�^�^���������������؂���
'�y�@���@���@�z�FJava�̃f�[�^�^
'�y�@�߂�l�@�z�FTrue:�������AFalse:�������Ȃ�
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
'�y�@�@�@�\�@�z�F�����̃e�[�u���̎�L�[���������̂��߂̃V�[�P���X�����擾����
'�y�@���@���@�z�F�X�L�[�}��, �e�[�u����
'�y�@�߂�l�@�z�F�V�[�P���X��
'********************************************************************************
Private Function GetSequenceName(ByVal strSchemaName, ByVal strTableName)
	Dim result
	result = ""
	
	'�Y���e�[�u���̎�L�[���������̂��߂̃V�[�P���X�����Aini �t�@�C���Ɏw�肳��Ă��Ȃ����m�F
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
			Call WriteLogWarn("script.ini�Ŏw�肳�ꂽ�V�[�P���X(" & strSequenceName & ")�͑��݂��܂���B")
		End If
	End If
	
	GetSequenceName = result
End Function

'********************************************************************************
'�y�@�@�@�\�@�z�F������Oracle �V�[�P���X�����݂��邩���؂���
'�y�@���@���@�z�FOracle �V�[�P���X��
'�y�@�߂�l�@�z�FTrue:���݂���AFalse:���݂��Ȃ�
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
		Call WriteLogError(Err, "SQL���s���G���[�i" & strSql & "�j�B oo4o: " & objOracleManager.GetOracleErrorCode & " " & objOracleManager.GetOracleErrorMessage)
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
'�y�@�@�@�\�@�z�F�����̃e�[�u���A�J�����̊֘A��`�o�͕�������擾����
'�y�@���@���@�z�F�X�L�[�}��, �e�[�u����, �J������
'�y�@�߂�l�@�z�F�֘A��`�o�͕�����
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

	'�Y���J�����̊֘A��`(�֘A�̏��L�ґ��̒�`)���Aini �t�@�C���Ɏw�肳��Ă��Ȃ����m�F
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
				Call WriteLogWarn("script.ini�Ŏw�肳�ꂽ�֘A��`(" & strKey & ")�̌`��������Ă��܂��B1:1�A*:1�̂����ꂩ�Ŏw�肵�Ă��������B")
			End If
		Next
	End If

	'�Y���J�����̊֘A��`(�֘A�̔폊�L�ґ��̒�`)���Aini �t�@�C���Ɏw�肳��Ă��Ȃ����m�F
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
				Call WriteLogWarn("script.ini�Ŏw�肳�ꂽ�֘A��`(" & strKey & ")�̌`��������Ă��܂��B1:1�A*:1�̂����ꂩ�Ŏw�肵�Ă��������B")
			End If
		Next
	End If

	GetRelationshipText = result
End Function

'********************************************************************************
'�y�@�@�@�\�@�z�F�����̃e�[�u���A�J�����̔C�ӂ�
'                �t�B�[���h�A�m�e�[�V�����o�͕�������擾����
'�y�@���@���@�z�F�X�L�[�}��, �e�[�u����, �J������
'�y�@�߂�l�@�z�F�t�B�[���h�A�m�e�[�V�����o�͕�����
'********************************************************************************
Private Function GetFieldAnnotations(ByVal strSchemaName, ByVal strTableName, ByVal strColumnName)
	On Error Resume Next
	Dim result()

	If strSchemaName = "" Then
		strSchemaName = objIniFileReader.GetIniValue("DataBaseInfo", "Oracle_User", "")
	End If

	'�Y���J�����̔C�ӂ̃t�B�[���h�A�m�e�[�V�������Aini �t�@�C���Ɏw�肳��Ă��Ȃ����m�F
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
					'�p�b�P�[�W�w�肪�Ȃ�(�A�m�e�[�V����������p�b�P�[�W)�̏ꍇ
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
					'�p�b�P�[�W�w�肪�Ȃ�(�A�m�e�[�V����������p�b�P�[�W)�̏ꍇ
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
'�y�@�@�@�\�@�z�F�G���e�B�e�B��1��������
'�y�@���@���@�z�F�X�L�[�}��, �e�[�u����
'�y�@�߂�l�@�z�F�Ȃ�
'********************************************************************************
Private Sub GenerateOneEntity(ByVal schemaName, ByVal tableName)
	On Error Resume Next
	Dim strSql
	Dim objRecordSet

	' �e�[�u���R�����g�擾
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
        Call WriteLogError(Err, "SQL���s���G���[�i" & strSql & "�j�B oo4o: " & objOracleManager.GetOracleErrorCode & " " & objOracleManager.GetOracleErrorMessage)
		Exit Sub
    End If
    
    If objRecordSet.EOF Then
    	If schemaName <> "" Then
    		WScript.Echo("�e�[�u�� " & schemaName & "." & tableName & " �͑��݂��Ȃ�������������܂���B")
    	Else
    		WScript.Echo("�e�[�u�� " & tableName & " �͑��݂��܂���B")
    	End If
		Exit Sub
    End If
    
    Dim strTableComment
    strTableComment = objUtil.GetDef(objRecordSet.Fields(0).Value, "")
    Set objRecordSet = Nothing
    
    ' �J�������擾
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
        Call WriteLogError(Err, "SQL���s���G���[�i" & strSql & "�j�B oo4o: " & objOracleManager.GetOracleErrorCode & " " & objOracleManager.GetOracleErrorMessage)
		Exit Sub
    End If

	'��ۃN���X�����ɑ��݂��邩�m�F
	'���݂���ꍇ�́A��`����Ă���t�B�[���h��z��Ɋi�[
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

	'�e�[�u���ɑ��݂��Ȃ��t�B�[���h����ۃN���X�ɒ�`����Ă��Ȃ����m�F
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
						'�e�[�u���ɑ��݂��Ȃ��t�B�[���h���i�����ΏۊO�Ƃ��Ďw�肳��Ă��邩�m�F
						Set objClassFile = objFso.OpenTextFile("dist\" & strClassName & ".java", ForReading)
						strPreLine = ""
						strLine = ""
						Do While objClassFile.AtEndOfStream = False
							strPreLine = strLine
							strLine = objClassFile.ReadLine
							If strLine = f Then
								If Instr(strPreLine, "@Transient") = 0 Then
									WriteLogWarn(strClassName & ".java �Ńe�[�u���Ŗ���`�̃t�B�[���h��Transient�w��Ȃ��ɒ�`����Ă��܂�(" & f & ")")
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

    ' �v���C�}���L�[���擾
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
        Call WriteLogError(Err, "SQL���s���G���[�i" & strSql & "�j�B oo4o: " & objOracleManager.GetOracleErrorCode & " " & objOracleManager.GetOracleErrorMessage)
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

	'�G���e�B�e�B����(���ۃN���X)
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
			
			' �֘A��`import���o��
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
			strText = strText & " * " & strTmpString(0) & " Entity �̒��ۃN���X" & strNewLine
		ElseIf strLine = " * @version" Then
			strText = strText & strLine & " " & objUtil.FormatDate(Now, "YYYY/MM/DD") & strNewLine
		ElseIf Left(strLine, 27) = "public abstract class Clazz" Then
			strText = strText & Replace(strLine, "Clazz", strAbstractClassName) & strNewLine
		ElseIf Left(strLine, 1) = "}" Then
			For i = 0 To UBound(strColumnName)
				'�t�B�[���h�R�����g
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
				
				'�C�ӂ̃t�B�[���h�A�m�e�[�V�����o��
				strFieldAnnotations = GetFieldAnnotations(schemaName, tableName, strColumnName(i))
				If IsArray(strFieldAnnotations) Then
					For k = 0 To UBound(strFieldAnnotations, 2)
						strImport = strImport & strFieldAnnotations(0, k)
						strText = strText & vbTab & strFieldAnnotations(1, k)
					Next
				End If
				
				'�J��������Java�\��ꂾ�����ꍇ�� @Column �A�m�e�[�V����������
				If IsJavaReserved(strColumnName(i)) Then
					strImport = strImport & "import javax.persistence.Column;"
					strText = strText & vbTab & "@Column(name = """ & strColumnName(i) & """)" & strNewLine
				End If
				
				'�t�B�[���h�o��
				strText = strText & vbTab & "public " & DBTypeToJavaType(strColumnType(i), schemaName, tableName, strColumnName(i)) & " " & ColumnNameToFieldName(strColumnName(i)) & ";" & strNewLine & strNewLine
			Next
			
			' �֘A��`�o��
			If strRelationshipText <> "" Then
				strText = strText & vbTab & Replace(Replace(strRelationshipText, strNewLine, strNewLine & vbTab), vbTab & strNewLine, strNewLine)
				'������ vbTab ���폜
				strText = Left(strText, Len(strText) - 1)
			End If
			
			' �N���X�̏I���J�b�R "}" �o��
			strText = strText & strLine & strNewLine
		Else
			strText = strText & strLine & strNewLine
		End If
	Wend
	objTemplateFile.Close
	Set objTemplateFile = Nothing

	'�t�@�C���o��
	objUtil.OutputTextFile "dist\" & strAbstractClassName & ".java", strImport & strText, strCharset

	'�G���e�B�e�B����(��ۃN���X)
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

		'�t�@�C���o��
		objUtil.OutputTextFile "dist\" & strClassName & ".java", strText, strCharset
	End If
End Sub

'********************************************************************************
'�y�@�@�@�\�@�z�F�G���e�B�e�B�����ׂĐ�������
'�y�@���@���@�z�F�X�L�[�}��
'�y�@�߂�l�@�z�F�Ȃ�
'********************************************************************************
Private Sub GenerateAllEntity(ByVal schemaName)
	On Error Resume Next
	Dim strSql
	Dim objRecordSet

	' �e�[�u�����X�g�擾
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
        Call WriteLogError(Err, "SQL���s���G���[�i" & strSql & "�j�B oo4o: " & objOracleManager.GetOracleErrorCode & " " & objOracleManager.GetOracleErrorMessage)
		Exit Sub
    End If
    
    If objRecordSet.EOF Then
   		WScript.Echo("�Ώۂ̃e�[�u�������݂��܂���B")
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
