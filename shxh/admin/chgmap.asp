<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
id=Request.QueryString("id")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("BA_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�") then Response.Redirect "../error.asp?id=046"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from map where id="&id ,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
displace=rst("displace")
affair=trim(rst("affair"))
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel='stylesheet' href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false  background='<%=bgimage%>' bgcolor='<%=bgcolor%>'>
<p align=center>��ͼϵͳ</p><hr>
�¼���������Ϊ��<br>
1.msg ��Ϣ ����msg ��˵��ܲˣ���������ߵ��˴�ʱ����ʾ��Ϣ����˵����ܲ�"<br>
2.fight ������ ����fight �޽ܣ�������ߵ��˴��������޽ܣ���ȷ���޽ܴ����ڹ������<br>
3 exit û�в���������ߵ��˴����˳�<br>
<form action='upmap.asp' method=post id=form1 name=form1>
<table width='80%' align=center border=3>
<tr><td>ID(ֻ��)</td><td><input type=text name=id readonly value=<%=id%>></td></tr>
<tr><td>���ƶ�(�߼�)</td><td><input type=text name=displace value="<%=displace%>" size=5 maxlength=5></td></tr>
<tr><td>�¼�(�ı�)</td><td><input type=text name=affair value="<%=affair%>" size=50 maxlength=25></td></tr>
<tr align=center><td colspan=2><input type=submit  value=' �� �� '> <input type=button value='����' onclick=javascript:location.href='showmap.asp';></td></tr>
</table>
</form>