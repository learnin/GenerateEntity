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
'�y�@�@�@�\�@�z�Foo4o���g�p����Oracle�ڑ��Ǘ����s���N���X
'********************************************************************************
Class OracleConnectionManager
	Private objSession
	Private objConnection
	Private objIniFileReader

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�R���X�g���N�^
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Private Sub Class_Initialize
		Set objSession = Nothing
		Set objConnection = Nothing
		Set objIniFileReader = New IniFileReader
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�f�X�g���N�^
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Private Sub Class_Terminate
		On Error Resume Next
		Call CloseSession
		Set objIniFileReader = Nothing
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�Fini�t�@�C���p�X���Z�b�g
	'�y�@���@���@�z�Fini�t�@�C���p�X
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Public Sub SetIniFile(ByVal strIniFilePath)
		objIniFileReader.IniFilePath = strIniFilePath
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�Foo4o�Z�b�V�������擾����
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
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
	'�y�@�@�@�\�@�z�Fini�t�@�C�����ڑ������擾���Aoo4o�ڑ����s��
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
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
	'�y�@�@�@�\�@�z�Foo4o�G���[���Z�b�g���s��
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Public Sub ResetOracleError
		On Error Resume Next
        objConnection.LastServerErrReset
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�Foo4o�G���[�R�[�h���擾����
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�Foo4o�G���[�R�[�h
	'********************************************************************************
	Public Function GetOracleErrorCode
		On Error Resume Next
        GetOracleErrorCode = objConnection.LastServerErr
	End Function

	'********************************************************************************
	'�y�@�@�@�\�@�z�Foo4o�G���[���b�Z�[�W���擾����
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�Foo4o�G���[���b�Z�[�W
	'********************************************************************************
	Public Function GetOracleErrorMessage
		On Error Resume Next
        GetOracleErrorMessage = objConnection.LastServerErrText
	End Function

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�g�����U�N�V�������J�n����
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Public Sub BeginTrans
		On Error Resume Next
        objSession.BeginTrans
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�g�����U�N�V�������R�~�b�g����
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Public Sub CommitTrans
		On Error Resume Next
        objSession.CommitTrans
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�g�����U�N�V���������[���o�b�N����
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Public Sub Rollback
		On Error Resume Next
        objSession.Rollback
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�X�V�nSQL�����s����
	'�y�@���@���@�z�FSQL
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Public Sub ExecuteSQL(ByVal strSql)
		On Error Resume Next
        objConnection.ExecuteSQL(strSql)
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�Q�ƌnSQL�����s����
	'�y�@���@���@�z�FSQL
	'�y�@�߂�l�@�z�F���ʃZ�b�g
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
	'�y�@�@�@�\�@�z�Foo4o�Z�b�V�����A�ڑ������
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Public Sub CloseSession
		On Error Resume Next
		
        Set objConnection = Nothing
		Set objSession = Nothing
	End Sub
End Class

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
