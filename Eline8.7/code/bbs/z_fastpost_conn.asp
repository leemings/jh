<%
dim dbfastpost,connfastpost,fastpoststr
dbfastpost="data/e_fastpost.asp"   '��ĳ����Լ������ݿ��ļ���
Set connfastpost = Server.CreateObject("ADODB.Connection")
fastpoststr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbfastpost)
connfastpost.Open fastpoststr
%>