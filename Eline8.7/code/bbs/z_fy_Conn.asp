<%
dim dbfy,connfy,fystr
dbfy="data/e_fyjy.asp"   '��ĳ����Լ������ݿ��ļ���
Set connfy = Server.CreateObject("ADODB.Connection")
fystr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbfy)
connfy.Open fystr
%>
