<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=Request.QueryString("un")
if instr(un,"'")<>0 or instr(un," ")<>0 or un="" then Response.Redirect "../error.asp?id=024"
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject ("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select * from �û� where ����='"&username&"' and ����>499"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=023"
rst.Close
sqlstr="select * from �û� where ����='"&un&"' and ״̬='����'"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=025"
rst.Close
set rst=nothing
conn.Execute "update �û� set ����=����-500 where ����='"&username&"'"
conn.Execute "update �û� set ״̬='����',����¼ʱ��=now() where ����='"&un&"'"
conn.Close 
set conn=nothing
%><head><title><%=Application("Ba_jxqy_systemname")%></title><LINK href="../chatroom/css.css" rel=stylesheet></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("Ba_jxqy_systemname")%>����</b></font></p>
<p align=center>�����Ѿ��ͷ���<%=un%>,Ҳлл���500���׻���������</p>
<p align=center>
  <input type="button" value=" �� �� " onClick="javascript:top.window.close();" name="button">
</p>
</body>