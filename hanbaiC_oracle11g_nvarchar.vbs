' ************************************************
' �Ǘ��Ҍ����ŃR�}���h�v�����v�g�Ŏ��s���������āA
' �Ō�� pause ����
' ************************************************
Set Shell = CreateObject("Shell.Application")
Set WshShell = Wscript.CreateObject("WScript.Shell")
if Wscript.Arguments.Count = 0 then
	ScriptFullName = WScript.ScriptFullName
	Shell.ShellExecute "cmd.exe", "/c cscript.exe """ & ScriptFullName & """ dummy_param & pause", "", "runas", 1
	WScript.Quit
end if

' ************************************************
' ��{�ݒ�
' Microsoft ODBC for Oracle �Ŏ��s�ł��܂�
' DSN ���쐬���ē���m�F���Ďw�肵�ĉ�����
' ( �Q�l:http://lightbox.matrix.jp/ginpro/patio.cgi?mode=view&no=225&type=ref )
' ************************************************
' ���̃X�N���v�g�����݂���f�B���N�g�����擾
strCurDir = WScript.ScriptFullName
strCurDir = Replace( strCurDir, WScript.ScriptName, "" )
strMdbPath = strCurDir & "�̔��Ǘ�C.mdb"

' Oracle �̃z�X�g������
' ( ���[�J���E�l�b�g�E�T�[�r�X�� )
strTarget = "{Oracle in instantclient_21_6}"	' ODBC �h���C�o
strDBQ = "localhost:1521/XEPDB1"	' �l�b�g�E�T�[�r�X���Ƃ��� XE �݂̂ł� OK
' �X�L�[�}(���[�U)
strSc = "LIGHTBOX02"
' �p�X���[�h
strPwd = InputBox("�p�X���[�h����͂��ĉ�����")

strDummy = "DUMMY" & Replace(Date,"/","") & Replace(Time,":","")

strMessage = "�Ώ� MDB �� " & strMdbPath & "�ł�" & vbCrLf & vbCrLf

strMessage = strMessage & "�� Oracle�̊��ł�" & vbCrLf
strMessage = strMessage & "ODBC �h���C�o : " & strTarget & vbCrLf
strMessage = strMessage & "�C���X�^���X : " & strDBQ & vbCrLf
strMessage = strMessage & "USER(�X�L�[�}) : " & strSc & vbCrLf
strMessage = strMessage & "PASS : " & strPwd & vbCrLf & vbCrLf
strMessage = strMessage & "�ꎞ�e�[�u�� : " & strDummy & vbCrLf & vbCrLf

strMessage = strMessage & "���Ƀe�[�u�������݂���ꍇ�̓��b�Z�[�W���o�܂���" & vbCrLf
strMessage = strMessage & "����ȊO�ł̓G���[���b�Z�[�W���o�܂����A��肠��܂���"
if vbCancel = MsgBox( strMessage, vbOkCancel ) then
	Wscript.Quit
end if

' ************************************************
' �����p������ݒ�
' ************************************************
' MDB �̐ڑ�������
strConnectMdb = _
"Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & strMdbPath & ";"

' Microsoft �� Oracle �p ODBC �h���C�o�̐ڑ������� (1)
strConnectOracle = _
"[ODBC;Driver=" & strTarget & ";DBQ=" & strDBQ &";UID=" & strSc & ";PWD=" & strPwd & "]"

' Microsoft �� Oracle �p ODBC �h���C�o�̐ڑ������� (2)
strConnectOracle2 = _
"Provider=MSDASQL;Driver=" & strTarget & ";DBQ=" & strDBQ &";UID=" & strSc & ";PWD=" & strPwd

' ************************************************
' ��������
' ************************************************
Set Cn = CreateObject("ADODB.Connection")
Set Cn2 = CreateObject("ADODB.Connection")
Cn.Open strConnectMdb
Cn2.Open strConnectOracle2

' ************************************************
' �R�[�h���̃}�X�^
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	�敪 NUMBER(4,0)" & _
"	,�R�[�h NVARCHAR2(10)" & _
"	,���� NVARCHAR2(50)" & _
"	,���l1 NUMBER(8,0)" & _
"	,���l2 NUMBER" & _
"	,�쐬�� DATE" & _
"	,�X�V�� DATE" & _
"	,primary key(�敪,�R�[�h)" & _
")"
Call OracleTransfer( "�R�[�h���̃}�X�^", "[�R�[�h���̃}�X�^]", Query )

' ************************************************
' �R���g���[���}�X�^
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	�L�[ NVARCHAR2(1)" & _
"	,������t DATE" & _
"	,����`�[ NUMBER(8,0)" & _
"	,��Ж� NVARCHAR2(50)" & _
"	,�g�D�R�[�h NVARCHAR2(4)" & _
"	,�N�Z�� NUMBER(2,0)" & _
"	,primary key(�L�[)" & _
")"
Call OracleTransfer( "�R���g���[���}�X�^","[�R���g���[���}�X�^]", Query )

' ************************************************
' ���b�Z�[�W�}�X�^
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	�R�[�h NVARCHAR2(4)" & _
"	,���b�Z�[�W NVARCHAR2(100)" & _
"	,primary key(�R�[�h)" & _
")"
Call OracleTransfer( "���b�Z�[�W�}�X�^","[���b�Z�[�W�}�X�^]", Query )

' ************************************************
' ����f�[�^
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	����敪 NVARCHAR2(2)" & _
"	,�`�[�ԍ� NUMBER(8,0)" & _
"	,�s NUMBER(2,0)" & _
"	,������t DATE" & _
"	,�����R�[�h NVARCHAR2(4)" & _
"	,���i�R�[�h NVARCHAR2(4)" & _
"	,���� NUMBER" & _
"	,�P�� NUMBER" & _
"	,���z NUMBER" & _
"	,�X�V�� NVARCHAR2(1)" & _
"	,primary key(����敪,�`�[�ԍ�,�s)" & _
")"
Call OracleTransfer( "����f�[�^","[����f�[�^]", Query )

' ************************************************
' ���i�}�X�^
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	���i�R�[�h NVARCHAR2(4)" & _
"	,���i�� NVARCHAR2(50)" & _
"	,�݌ɕ]���P�� NUMBER" & _
"	,�̔��P�� NUMBER" & _
"	,���i���� NVARCHAR2(3)" & _
"	,���i�敪 NVARCHAR2(1)" & _
"	,�쐬�� DATE" & _
"	,�X�V�� DATE" & _
"	,���l NVARCHAR2(2000)" & _
"	,�폜�t���O NVARCHAR2(1)" & _
"	,primary key(���i�R�[�h)" & _
")"
Call OracleTransfer( "���i�}�X�^","[���i�}�X�^]", Query )

' ************************************************
' ���i���ރ}�X�^
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	���i���� NVARCHAR2(3)" & _
"	,���� NVARCHAR2(50)" & _
"	,�쐬�� DATE" & _
"	,�X�V�� DATE" & _
"	,primary key(���i����)" & _
")"
Call OracleTransfer( "���i���ރ}�X�^","[���i���ރ}�X�^]", Query )

' ************************************************
' ���Ӑ�}�X�^
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	���Ӑ�R�[�h NVARCHAR2(4)" & _
"	,���Ӑ於 NVARCHAR2(50)" & _
"	,���Ӑ�敪 NVARCHAR2(1)" & _
"	,�S���� NVARCHAR2(4)" & _
"	,�X�֔ԍ� NVARCHAR2(7)" & _
"	,�Z���P NVARCHAR2(100)" & _
"	,�Z���Q NVARCHAR2(100)" & _
"	,�쐬�� DATE" & _
"	,�X�V�� DATE" & _
"	,���� NUMBER(2,0)" & _
"	,�����敪 NUMBER(1,0)" & _
"	,�x���� NUMBER(2,0)" & _
"	,���l NVARCHAR2(100)" & _
"	,primary key(���Ӑ�R�[�h)" & _
")"
Call OracleTransfer( "���Ӑ�}�X�^","[���Ӑ�}�X�^]", Query )

' ************************************************
' �Ј��}�X�^
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	�Ј��R�[�h NVARCHAR2(4)" & _
"	,���� NVARCHAR2(50)" & _
"	,�t���K�i NVARCHAR2(50)" & _
"	,���� NVARCHAR2(4)" & _
"	,���� NUMBER(1,0)" & _
"	,�쐬�� DATE" & _
"	,�X�V�� DATE" & _
"	,���^ NUMBER" & _
"	,�蓖 NUMBER" & _
"	,�Ǘ��� NVARCHAR2(4)" & _
"	,���N���� DATE" & _
"	,primary key(�Ј��R�[�h)" & _
")"
Call OracleTransfer( "�Ј��}�X�^","[�Ј��}�X�^]", Query )

' ************************************************
' �X�֔ԍ��}�X�^
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	�X�֔ԍ� NVARCHAR2(7)" & _
"	,�s���{�����J�i NVARCHAR2(255)" & _
"	,�s�撬�����J�i NVARCHAR2(255)" & _
"	,���於�J�i NVARCHAR2(255)" & _
"	,�s���{���� NVARCHAR2(255)" & _
"	,�s�撬���� NVARCHAR2(255)" & _
"	,���於 NVARCHAR2(255)" & _
")"
Call OracleTransfer( "�X�֔ԍ��}�X�^","[�X�֔ԍ��}�X�^]", Query )

' ************************************************
' �����\��f�[�^
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	���Ӑ�R�[�h NVARCHAR2(4)" & _
"	,�x���� DATE" & _
"	,�`�[���v���z NUMBER(10,0)" & _
"	,�`�[�ԍ� NUMBER(10,0)" & _
")"
Call OracleTransfer( "�����\��f�[�^","[�����\��f�[�^]", Query )

' ************************************************
' ���i�W�v
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	���i�R�[�h NVARCHAR2(4)" & _
"	,�o�ߌ� NUMBER(2,0)" & _
"	,�������㐔�� NUMBER(10,0)" & _
"	,����������z NUMBER(10,0)" & _
"	,�X�V�� DATE" & _
"	,�g�D�R�[�h NVARCHAR2(4)" & _
"	,primary key(���i�R�[�h,�o�ߌ�)" & _
")"
Call OracleTransfer( "���i�W�v","[���i�W�v]", Query )

' ************************************************
' ���Ӑ�W�v
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	������ NVARCHAR2(4)" & _
"	,�o�ߌ� NUMBER(2,0)" & _
"	,����������z NUMBER(10,0)" & _
"	,�X�V�� DATE" & _
"	,�g�D�R�[�h NVARCHAR2(4)" & _
"	,primary key(������,�o�ߌ�)" & _
")"
Call OracleTransfer( "���Ӑ�W�v","[���Ӑ�W�v]", Query )

' ************************************************
' �Ј��ύX����
' ************************************************
Query = _
"create table " & strDummy & " (" & _
"	�Ј��R�[�h NVARCHAR2(4)" & _
"	,���� NVARCHAR2(50)" & _
"	,�t���K�i NVARCHAR2(50)" & _
"	,���� NVARCHAR2(4)" & _
"	,���� NUMBER(1,0)" & _
"	,�쐬�� DATE" & _
"	,�X�V�� DATE" & _
"	,���^ NUMBER" & _
"	,�蓖 NUMBER" & _
"	,�Ǘ��� NVARCHAR2(4)" & _
"	,���N���� DATE" & _
")"
Call OracleTransfer( "�Ј��ύX����","[�Ј��ύX����]", Query )

' ************************************************
' �r���[
' ************************************************
Query = _
"create or replace view V_���i�ꗗ as" & _
"	SELECT ���i�}�X�^.���i�R�[�h" & _
"	, ���i�}�X�^.���i��" & _
"	, ���i�}�X�^.�̔��P��" & _
"	, ���i���ރ}�X�^.���i����" & _
"	, ���i���ރ}�X�^.���� AS ���ޖ�" & _
"	, ���i�}�X�^.���i�敪" & _
"	, �R�[�h���̃}�X�^.���� AS �敪��" & _
" from" & _
"	(���i�}�X�^ LEFT JOIN ���i���ރ}�X�^" & _
"	ON ���i�}�X�^.���i���� = ���i���ރ}�X�^.���i����" & _
"	) LEFT JOIN �R�[�h���̃}�X�^" & _
"	ON ���i�}�X�^.���i�敪 = �R�[�h���̃}�X�^.�R�[�h" & _
" where" & _
"	�R�[�h���̃}�X�^.�敪 = 3 and �폜�t���O is NULL"
RunOracle( Query )

Query = _
"create or replace view V_������t as" & _
"	SELECT �R���g���[���}�X�^.������t" & _
"	FROM �R���g���[���}�X�^" & _
"	WHERE �R���g���[���}�X�^.�L�[ = '1'"
RunOracle( Query )

Query = _
"create or replace view V_���Ӑ�䒠 as" & _
"	SELECT ����f�[�^.�����R�[�h" & _
"	, ���Ӑ�}�X�^.���Ӑ於" & _
"	, ����f�[�^.������t" & _
"	, ����f�[�^.����敪" & _
"	, ����f�[�^.�`�[�ԍ�" & _
"	, ����f�[�^.�s" & _
"	, ����f�[�^.���i�R�[�h" & _
"	, ���i�}�X�^.���i��" & _
"	, ����f�[�^.����" & _
"	, ����f�[�^.�P��" & _
"	, ����f�[�^.���z" & _
" from" & _
"	(����f�[�^ INNER JOIN ���i�}�X�^" & _
"	ON ����f�[�^.���i�R�[�h=���i�}�X�^.���i�R�[�h" & _
"	) INNER JOIN ���Ӑ�}�X�^" & _
"	ON ����f�[�^.�����R�[�h=���Ӑ�}�X�^.���Ӑ�R�[�h" & _
" where" & _
"	����f�[�^.����敪 = '10'"
RunOracle( Query )

Query = _
"create or replace view V_�Ј��ꗗ as" & _
" select �Ј��R�[�h" & _
"	,����" & _
"	,�t���K�i" & _
"	,����1.���� as ����" & _
"	,����" & _
"	,����2.���� as ������" & _
" from �Ј��}�X�^" & _
"	,�R�[�h���̃}�X�^ ����1" & _
"	,�R�[�h���̃}�X�^ ����2" & _
" where to_char(����) = ����1.�R�[�h" & _
"   and ����1.�敪 = 1" & _
"   and ���� = ����2.�R�[�h" & _
"   and ����2.�敪 = 2"
RunOracle( Query )

Query = _
"create or replace view" & _
"	PROC_ERROR" & _
" as" & _
" select * " & _
" from USER_ERRORS"
RunOracle( Query )

Query = _
"create or replace view" & _
"	PROC_LIST" & _
" as" & _
" select OBJECT_NAME as ""�v���V�[�W����"" " & _
"	,STATUS as ""���""  " & _
"	,OBJECT_TYPE  as ""�^�C�v"" " & _
"	,CREATED  as ""�쐬��"" " & _
"	,LAST_DDL_TIME  as ""�X�V��"" " & _
" from USER_OBJECTS " & _
" where OBJECT_TYPE in ('FUNCTION','PROCEDURE') "
RunOracle( Query )

Query = _
"create or replace view" & _
"	PROC_TEXT" & _
" as" & _
" select * from USER_SOURCE"
RunOracle( Query )

' ************************************************
' �I��
' ************************************************

Cn2.Close
Cn.Close

Wscript.Echo "�������I�����܂���"

' ************************************************
' Oracle �]��
' ************************************************
function OracleTransfer( strTarget, strTable, QueryCreate )

	Dim Query

	Query = "drop table " & strTarget
	RunOracle( Query )

	RunOracle( QueryCreate )

	Query = "insert into " & strConnectOracle & "." & strDummy & _
	" select * from " & strTable

	RunMdb( Query )

	Query = "alter table " & strDummy & " rename to " & strTarget
	RunOracle( Query )

end function

' ************************************************
' MDB ���s
' ************************************************
function RunMdb( Query )

	on error resume next
	Cn.Execute Query
	if Err.Number <> 0then
		Wscript.Echo Err.Description & vbCrLf & Query
	end if
	on error goto 0

end function

' ************************************************
' Oracle ���s
' ************************************************
function RunOracle( Query )

	on error resume next
	Cn2.Execute Query
	if Err.Number <> 0then
		Wscript.Echo Err.Description & vbCrLf & Query
	end if
	on error goto 0

end function
