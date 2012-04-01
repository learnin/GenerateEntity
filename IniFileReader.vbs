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
	'�y�@�@�@�\�@�z�F�R���X�g���N�^
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Private Sub Class_Initialize
		Set objUtil = New Util
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�f�X�g���N�^
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Private Sub Class_Terminate
		Set objUtil = Nothing
	End Sub

	'********************************************************************************
	'�v���p�e�B��`
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
	'�y�@�@�@�\�@�z�F�����̕�������Z�N�V�������̏����ɂ��ĕԂ�
	'�y�@���@���@�z�FstrSectionName : �Z�N�V������
	'�y�@�߂�l�@�z�F�����̕�����̐擪�A�����ɂ��ꂼ��"["�A"]"������������
	'********************************************************************************
	Private Function GetSection(ByVal strSectionName)
		GetSection = "[" & Trim(strSectionName) & "]"
	End Function

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�����̕����񂩂� key ��Ԃ�
	'�y�@���@���@�z�FstrString : key=value �`���̕�����
	'�y�@�߂�l�@�z�Fkey �ƂȂ镶����
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
	'�y�@�@�@�\�@�z�F�����̕����񂩂� value ��Ԃ�
	'�y�@���@���@�z�FstrString : key=value �`���̕�����
	'�y�@�߂�l�@�z�Fvalue �ƂȂ镶����
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
	'�y�@�@�@�\�@�z�Fini�t�@�C����������̃Z�N�V�����Akey �ɊY������ value ��Ԃ�
	'�y�@���@���@�z�FstrSection : �Z�N�V������
	'�@�@�@�@�@�@�@�FstrKey : key
	'              �FstrDefaultValue : �f�t�H���g�l
	'�y�@�߂�l�@�z�F�����̃Z�N�V�����Akey �ɊY������ value�B
	'                value ���擾�ł��Ȃ������ꍇ�� strDefaultValue�B
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
	'�y�@�@�@�\�@�z�Fini�t�@�C����������̃Z�N�V�����A
	'�@�@�@�@�@�@�@�@������ key �ƑO����v���� key �ƁA�Y������ value ��2�����z���Ԃ�
	'�y�@���@���@�z�FstrSection : �Z�N�V������
	'�@�@�@�@�@�@�@�FstrKey : key
	'�y�@�߂�l�@�z�F�����̃Z�N�V�����Akey �ƑO����v���� key �ƁA�Y������ value ��2�����z��B
	'                value ���擾�ł��Ȃ������ꍇ�� ""�B�������A�����IsArray�ōs���Ă��������B
	'�@�@�@�@�@�@�@�@GetIniValueByKeyStartsWith(1, n)
	'�@�@�@�@�@�@�@�@�@GetIniValueByKeyStartsWith(0, n)�Fn�ڂ̃L�[
	'�@�@�@�@�@�@�@�@�@GetIniValueByKeyStartsWith(1, n)�Fn�ڂ̒l
	'********************************************************************************************
	Public Function GetIniValueByKeyStartsWith(ByVal strSection, ByVal strKey)
		On Error Resume Next
		GetIniValueByKeyStartsWith = GetIniValueByKeyStartsOrEndsWith(strSection, strKey, True)
	End Function

	'********************************************************************************************
	'�y�@�@�@�\�@�z�Fini�t�@�C����������̃Z�N�V�����A
	'�@�@�@�@�@�@�@�@������ key �ƌ����v���� key �ƁA�Y������ value ��2�����z���Ԃ�
	'�y�@���@���@�z�FstrSection : �Z�N�V������
	'�@�@�@�@�@�@�@�FstrKey : key
	'�y�@�߂�l�@�z�F�����̃Z�N�V�����Akey �ƌ����v���� key �ƁA�Y������ value ��2�����z��B
	'                value ���擾�ł��Ȃ������ꍇ�� ""�B�������A�����IsArray�ōs���Ă��������B
	'�@�@�@�@�@�@�@�@GetIniValueByKeyEndsWith(1, n)
	'�@�@�@�@�@�@�@�@�@GetIniValueByKeyEndsWith(0, n)�Fn�ڂ̃L�[
	'�@�@�@�@�@�@�@�@�@GetIniValueByKeyEndsWith(1, n)�Fn�ڂ̒l
	'********************************************************************************************
	Public Function GetIniValueByKeyEndsWith(ByVal strSection, ByVal strKey)
		On Error Resume Next
		GetIniValueByKeyEndsWith = GetIniValueByKeyStartsOrEndsWith(strSection, strKey, False)
	End Function

	'********************************************************************************************
	'�y�@�@�@�\�@�z�Fini�t�@�C����������̃Z�N�V�����A
	'�@�@�@�@�@�@�@�@������ key �ƑO����v�܂��͌����v���� key �ƁA�Y������ value ��2�����z���Ԃ�
	'�y�@���@���@�z�FstrSection : �Z�N�V������
	'�@�@�@�@�@�@�@�FstrKey : key
	'              �FbIsStarts : True:�O����v False:�����v
	'�y�@�߂�l�@�z�F�����̃Z�N�V�����Akey �ƑO����v�܂��͌����v���� key �ƁA�Y������ value ��2�����z��B
	'                value ���擾�ł��Ȃ������ꍇ�� ""�B�������A�����IsArray�ōs���Ă��������B
	'�@�@�@�@�@�@�@�@GetIniValueByKeyEndsWith(1, n)
	'�@�@�@�@�@�@�@�@�@GetIniValueByKeyEndsWith(0, n)�Fn�ڂ̃L�[
	'�@�@�@�@�@�@�@�@�@GetIniValueByKeyEndsWith(1, n)�Fn�ڂ̒l
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
'�y�@�@�@�\�@�z�F�����̃t�@�C�����C���N���[�h����v���V�[�W��
'�y�@���@���@�z�FstrFilePath : �C���N���[�h����VBScript�t�@�C����
'�y�@�߂�l�@�z�F�Ȃ�
'********************************************************************************
Sub Include(ByVal strFilePath)
	On Error Resume Next

	Dim objFso, objTextStream

    Set objFso = WScript.CreateObject("Scripting.FileSystemObject")

	'�ǂݎ���p�Ńt�@�C�����J��
	Set objTextStream = objFso.OpenTextFile(strFilePath, 1, False)

	'�t�@�C���̓��e�����s
	ExecuteGlobal(objTextStream.ReadAll)

    objTextStream.Close
	Set objTextStream = Nothing
	Set objFso = Nothing
End Sub
