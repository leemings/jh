<%
dim dbplus,connplus,plusstr
dbplus="data/e_plus.asp"   '��ĳ����Լ������ݿ��ļ���
Set connplus = Server.CreateObject("ADODB.Connection")
plusstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbplus)
connplus.Open plusstr
%>