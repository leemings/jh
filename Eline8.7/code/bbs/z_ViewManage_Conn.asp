<%
dim dbvm,connvm,vmstr
dbvm="data/e_ViewManage.asp"   '��ĳ����Լ������ݿ��ļ���
Set connvm = Server.CreateObject("ADODB.Connection")
vmstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbvm)
connvm.Open vmstr
%>
