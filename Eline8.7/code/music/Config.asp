<!-- #include file="conn.asp" -->
<!-- #include file="inc/function.asp" -->
<%
dim Title_Name,CategoryName,CategoryName_CHS,username,URL,kicktime

'��Ŀ������
CategoryName="SoftDown"   '��Ҫ�Ķ�

'��վ����
Title_Name="һ������ ==>> " '���ԸĶ�

'Ƶ������
CategoryName_CHS="E�߽���" '���ԸĶ�

UserName= session("sUserName") '��Ҫ�Ķ�

'���л���
SystemType="Windows.Net,WinXP,Win2000,NT,WinME,Win9X,Dos,Linux,Unix,Mac,ASP����,PHP����,CGI����"  '���ԸĶ�

'������ʱĬ�����л���
DefaultSystemType="WinXP,Win2000,NT,WinME,Win9X" '���ԸĶ�

'����ɾ���������ݱ��ж��ٷ����ڲ���û�,��λΪ����

kicktime=2400

URL="http://www.198962.com" '������Ŀ��ַ ���ԸĶ�

FolderPath="../Down/"  '��̬��ҳ����ļ��У�����ڹ�������ļ��е�λ�á�

%>