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
'�y�@�@�@�\�@�z�F���[�e�B���e�B�N���X
'********************************************************************************
Class Util

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�R���X�g���N�^
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	private Sub Class_Initialize
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�f�X�g���N�^
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	private Sub Class_Terminate
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�����Ŏw�肳�ꂽ�t�H�[�}�b�g�̓������擾
	'�y�@���@���@�z�F���t�I�u�W�F�N�g
	'                �t�H�[�}�b�g������
	'�y�@�߂�l�@�z�F�����Ŏw�肳�ꂽ�t�H�[�}�b�g�̓���
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
	'�y�@�@�@�\�@�z�F�w�肵�������񂪁ANull�l�܂���Empty�l�܂���
	'                ��(""�┼�p�E�S�p�X�y�[�X�̂�)�̏ꍇ�ɁA�����Ŏw�肳�ꂽ�������Ԃ�
	'�y�@���@���@�z�F���肷�镶����
	'                �f�t�H���g������
	'�y�@�߂�l�@�z�F������Null�܂���Empty�܂��͋󔒂̏ꍇ�͎w�肳�ꂽ������B
	'                ����ȊO�͂��̂܂�
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
	'�y�@�@�@�\�@�z�F�X�N���v�g�t�@�C���̃f�B���N�g���p�X���擾
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�X�N���v�g�t�@�C���̃f�B���N�g���p�X(�����Ɂu\�v�͂��Ȃ�)
	'********************************************************************************
	Public Function GetScriptDir
		GetScriptDir = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName) - 1)
	End Function

	'********************************************************************************
	'�y�@�@�@�\�@�z�FSQL�̃V���O���N�H�[�g���G�X�P�[�v����
	'�y�@���@���@�z�F������
	'�y�@�߂�l�@�z�F�G�X�P�[�v���ꂽ������
	'********************************************************************************
	Public Function SqlEscape(ByVal strVal)
		SqlEscape = "'" & Replace(strVal, "'", "''") & "'"
	End Function

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�w�肳�ꂽ�����R�[�h�Ńe�L�X�g�t�@�C���o�͂��s��
	'�y�@���@���@�z�F�o�̓t�@�C���t���p�X, �o�͓��e, �����R�[�h(ex. "utf-8")
	'�y�@�߂�l�@�z�F�Ȃ�
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
			'�W��API�ł�UTF-8�ɂ�BOM���o�͂���邽�߁A
			'�ŏ���3byte���X�L�b�v���ăo�C�i���ŃX�g���[����ǂݍ���ŏo�͂���
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
	'�y�@�@�@�\�@�z�F�����̕����񂩂�擪�̃X�y�[�X�܂��̓^�u�����������������Ԃ�
	'�y�@���@���@�z�FstrString : ������
	'�y�@�߂�l�@�z�F�����̕����񂩂�擪�̃X�y�[�X�܂��̓^�u����������������
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
	'�y�@�@�@�\�@�z�F�����̕����񂩂疖���̃X�y�[�X�܂��̓^�u�����������������Ԃ�
	'�y�@���@���@�z�FstrString : ������
	'�y�@�߂�l�@�z�F�����̕����񂩂疖���̃X�y�[�X�܂��̓^�u����������������
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
	'�y�@�@�@�\�@�z�F�����̕����񂩂�擪�Ɩ����̃X�y�[�X�܂��̓^�u�����������������Ԃ�
	'�y�@���@���@�z�FstrString : ������
	'�y�@�߂�l�@�z�F�����̕����񂩂�擪�Ɩ����̃X�y�[�X�܂��̓^�u����������������
	'********************************************************************************
	Public Function TrimSpaceTab(ByVal strString)
		On Error Resume Next
		TrimSpaceTab = RTrimSpaceTab(LTrimSpaceTab(strString))
	End Function

End Class
