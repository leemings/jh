<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�") then Response.Redirect "../error.asp?id=046"
un=Request("search")
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open application("Ba_jxqy_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
sqlstr="SELECT * FROM �û� where ����='"&un&"'"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then
	Response.Redirect "../error.asp?id=101"
else
%>

<head>
<link rel="stylesheet" href="../chatroom/css.css">
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=0 background="<%=bgimage%>" bgcolor="<%=bgcolor%>">

<form method=post action=updateuser.asp><table width=100% border=3><tr><td  width=30%>id(readonly)</td><td><input type=text readonly name=id value="<%=rst("id")%>"></td></tr>
<%	for i=1 to rst.Fields.Count-1
		select case rst.Fields(i).type
		case 202,130
			texttype="�ַ���"
			textmaxlength=rst.Fields(i).DefinedSize
			textsize=textmaxlength
			if textsize>50 then textsize=50
		case 3 
			texttype="������"
			textmaxlength=10
			textsize=10
		case 7
			texttype="����/ʱ��"
			textmaxlength=20
			textsize=20
		case 11
			texttype="�߼�"
			textmaxlength=5
			textsize=5
		case else
			texttype=rst.Fields(i).type
			textmaxlength=rst.Fields(i).ActualSize
			textsize=rst.Fields(i).ActualSize
		end select
%>
		<tr><td><%=rst.Fields(i).Name&"("&texttype&")"%></td><td><input type=text name="<%=rst.Fields(i).Name%>" value="<%=rst.Fields(i).value%>" size="<%=textsize%>" maxlength="<%=maxlength%>"></td></tr>
<%	next%>
	<tr align=center><td colspan=2 align=center><input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=un%>&ucur=ҩƷ';" value=' ҩ Ʒ '> <input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=un%>&ucur=����';" value=' �� �� '>  <input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=un%>&ucur=����';" value=' �� �� ' > <input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=un%>&ucur=ͷ��';" value=' ͷ �� ' > <input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=un%>&ucur=��Ʒ';" value=' �� Ʒ ' > <input type=submit value=' �� �� ' name=submit> <input type=submit value=' ɾ �� ' name=submit> <input type=reset value=' �� �� '> <input type="button" name="ok" value=" �� �� " onclick=javascript:history.go(-1)></td></tr>
	</table></form>
<%end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>