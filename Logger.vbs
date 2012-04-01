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

'���O���x����`
Const Logger_LOG_LEVEL_DEBUG = 1
Const Logger_LOG_LEVEL_INFO = 2
Const Logger_LOG_LEVEL_WARN = 3
Const Logger_LOG_LEVEL_ERROR = 4
Const Logger_LOG_LEVEL_FATAL = 5

'���O�o�͐��`
Const Logger_LOG_TYPE_FILE = 1
Const Logger_LOG_TYPE_EVENTLOG = 2

'********************************************************************************
'�y�@�@�@�\�@�z�F���O�Ǘ��N���X
'********************************************************************************
Class Logger
	private intLogLevel
	private intLogType
	private objWshShell
	Private objFso
	private strLogFilePath
	private objUtil

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�R���X�g���N�^
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	private Sub Class_Initialize
		intLogLevel = Logger_LOG_LEVEL_INFO
		intLogType = Logger_LOG_TYPE_FILE

		Set objWshShell = WScript.CreateObject("WScript.Shell")
		Set objFso = CreateObject("Scripting.FileSystemObject")
		Set objUtil = New Util
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�F�f�X�g���N�^
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	private Sub Class_Terminate
		On Error Resume Next
		Set objWshShell = Nothing
		Set objFso = Nothing
		Set objUtil = Nothing
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�F���O���x�����Z�b�g����
	'�y�@���@���@�z�F���O���x��������
	'�y�@�߂�l�@�z�F�Ȃ�
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
	'�y�@�@�@�\�@�z�F���O��ʂ��Z�b�g����
	'�y�@���@���@�z�F���O���
	'�y�@�߂�l�@�z�F�Ȃ�
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
	'�y�@�@�@�\�@�z�F���O�t�@�C���p�X���Z�b�g����
	'�y�@���@���@�z�F���O�t�@�C���t���p�X
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Public Sub SetLogFilePath(ByVal logFilePath)
		strLogFilePath = logFilePath
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�FDEBUG���x�����L�������肷��
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�L���Ȃ�True�A�����Ȃ�False
	'********************************************************************************
	Public Function IsDebugEnabled
		If intLogLevel <= Logger_LOG_LEVEL_DEBUG Then
			IsDebugEnabled = True
		Else
			IsDebugEnabled = False
		End If
	End Function

	'********************************************************************************
	'�y�@�@�@�\�@�z�FINFO���x�����L�������肷��
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�L���Ȃ�True�A�����Ȃ�False
	'********************************************************************************
	Public Function IsInfoEnabled
		If intLogLevel <= Logger_LOG_LEVEL_INFO Then
			IsInfoEnabled = True
		Else
			IsInfoEnabled = False
		End If
	End Function

	'********************************************************************************
	'�y�@�@�@�\�@�z�FWARN���x�����L�������肷��
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�L���Ȃ�True�A�����Ȃ�False
	'********************************************************************************
	Public Function IsWarnEnabled
		If intLogLevel <= Logger_LOG_LEVEL_WARN Then
			IsWarnEnabled = True
		Else
			IsWarnEnabled = False
		End If
	End Function

	'********************************************************************************
	'�y�@�@�@�\�@�z�FERROR���x�����L�������肷��
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�L���Ȃ�True�A�����Ȃ�False
	'********************************************************************************
	Public Function IsErrorEnabled
		If intLogLevel <= Logger_LOG_LEVEL_ERROR Then
			IsErrorEnabled = True
		Else
			IsErrorEnabled = False
		End If
	End Function

	'********************************************************************************
	'�y�@�@�@�\�@�z�FFATAL���x�����L�������肷��
	'�y�@���@���@�z�F�Ȃ�
	'�y�@�߂�l�@�z�F�L���Ȃ�True�A�����Ȃ�False
	'********************************************************************************
	Public Function IsFatalEnabled
		If intLogLevel <= Logger_LOG_LEVEL_FATAL Then
			IsFatalEnabled = True
		Else
			IsFatalEnabled = False
		End If
	End Function

	'********************************************************************************
	'�y�@�@�@�\�@�z�FDEBUG���O���o��
	'�y�@���@���@�z�F�G���[���b�Z�[�W
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Public Sub Debug(ByVal strMessage)
		On Error Resume Next
		If IsDebugEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_DEBUG, strMessage)
		End If
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�FINFO���O���o��
	'�y�@���@���@�z�F�G���[���b�Z�[�W
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Public Sub Info(ByVal strMessage)
		On Error Resume Next
		If IsInfoEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_INFO, strMessage)
		End If
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�FWARN���O���o��
	'�y�@���@���@�z�F�G���[���b�Z�[�W
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Public Sub Warn(ByVal strMessage)
		On Error Resume Next
		If IsWarnEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_WARN, strMessage)
		End If
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�FERROR���O���o��
	'�y�@���@���@�z�F�G���[���b�Z�[�W
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Public Sub Error(ByVal strMessage)
		On Error Resume Next
		If IsErrorEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_ERROR, strMessage)
		End If
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�FFATAL���O���o��
	'�y�@���@���@�z�F�G���[���b�Z�[�W
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Public Sub Fatal(ByVal strMessage)
		On Error Resume Next
		If IsFatalEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_FATAL, strMessage)
		End If
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�F���O��ʂɂ��w��̃��O���o��
	'�y�@���@���@�z�F���O���x��
	'                �G���[���b�Z�[�W
	'�y�@�߂�l�@�z�F�Ȃ�
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
	'�y�@�@�@�\�@�z�F�C�x���g���O���o��
	'�y�@���@���@�z�F���O���x��
	'                �G���[���b�Z�[�W
	'�y�@�߂�l�@�z�F�Ȃ�
	'********************************************************************************
	Private Sub WriteEventLog(ByVal logLevel, ByVal strMessage)
		On Error Resume Next

		Dim intEventLogLevel

		'Logger�̃��O���x���ƃC�x���g���O�̃��O���x���̃}�b�s���O
		If logLevel = Logger_LOG_LEVEL_DEBUG Then
			'���
			intEventLogLevel = 4
		ElseIf logLevel = Logger_LOG_LEVEL_INFO Then
			'���
			intEventLogLevel = 4
		ElseIf logLevel = Logger_LOG_LEVEL_WARN Then
			'�x��
			intEventLogLevel = 2
		ElseIf logLevel = Logger_LOG_LEVEL_ERROR Then
			'�G���[
			intEventLogLevel = 1
		ElseIf logLevel = Logger_LOG_LEVEL_FATAL Then
			'�G���[
			intEventLogLevel = 1
		End If

		objWshShell.LogEvent intEventLogLevel, strMessage
	End Sub

	'********************************************************************************
	'�y�@�@�@�\�@�z�F���O�t�@�C�����o��
	'�y�@���@���@�z�F���O���x��
	'                �G���[���b�Z�[�W
	'�y�@�߂�l�@�z�F�Ȃ�
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
