<%
dim conn,connstr,username,password,database,path
username="eline"			'���ݿ��½��
password="e_flash"			'���ݿ�����
database="eline_flash"			'���ݿ���
path="(local)"			'���ݿ����ڵ�IP��ַ������Ǳ������ݿ����ø���
set conn=Server.CreateObject("adodb.Connection")
connstr="PROVIDER=SQLOLEDB;DATA SOURCE="&path&";UID="&username&";PWD="&password&";DATABASE="&database
conn.open connstr
%>