<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ���� from �û� where ����='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
silver=rst("����")
rst.Close
set conn=nothing
bet=100
if silver<bet then silver=bet
response.write "<html><head><link rel=stylesheet href='../style.css'></head><body bgcolor='"&Application("Ba_jxqy_backgroundcolor")&"' background='"&Application("Ba_jxqy_backgroundimage")&"'><p align=center><form action='pdstart.asp' method=post >"&username&" BA$��"&silver&" ��ע��<input type=text maxlength=4 size=4 name=pcbet value='"&bet&"'> <input type=submit value=' �� ע '> <input type=button value=' �� �� ' onclick=javascript:parent.location.replace('dubo.asp');></form></p></body></html>"
%>

