<%
dim Database,CheckLogin,LoginURL,DatabaseMember,DatabaseUserName,DatabaseCash,BetLimit
CheckLogin=request.cookies("aspsky")("username")	'��ʲô��ʽ��¼��¼���������Ҫ��""
LoginURL="index.asp"				'��¼ҳ��ĵ�ַ
Database="data/eline_bbs_6.3.0.asp"				'���ݿ��ַ
DatabaseMember="user"		'�����û����ݵı���
DatabaseUserName="UserName"		'���ݿ����û������ֶ�����
DatabaseCash="userWealth"			'���ݿ��л��ҵ��ֶ�����
BetLimit=100				'ÿ��ˮ���������ע���
%>