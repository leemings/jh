<%
dim dbvm,connvm,vmstr
dbvm="data/e_ViewManage.asp"   '请改成您自己的数据库文件名
Set connvm = Server.CreateObject("ADODB.Connection")
vmstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbvm)
connvm.Open vmstr
%>
