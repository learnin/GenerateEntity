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

' �ǂݎ���p
Const FOR_READING = 1

' �t�B�[���h����Java�\���ɊY������ꍇ�ɂ���t�B�[���h���ړ���
Const FIELD_PREFIX_ON_JAVA_RESERVED = "a"

' Java�\���
Const JAVA_RESERVED = "CLASS,INTERFACE,PUBLIC,PROTECTED,PRIVATE,IMPORT,STATIC,ABSTRACT,FINAL,TRANSIENT,INT,BOOLEAN,DOUBLE,LONG,FLOAT,BYTE,EXTENDS,IMPLEMENTS"