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
id=Request.QueryString("id")
opt=Request.QueryString("opt")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
msg="<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false  background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center>ҩƷ����</p><hr><form action=modcur3.asp method=post><table border=3 width=80% align=center><tr align=center bgcolor=FFFF00><td>����</td><td>ֵ<input type=hidden name='schid' value='"&id&"'></td></tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �̵�  where id="&id,conn
if opt="����" then
	for i=1 to rst.Fields.Count-1
		if rst.Fields(i).type=202 then
			texttype="�ַ���"
			textmaxlength=rst.Fields(i).DefinedSize
			textsize=textmaxlength
			if textsize>50 then textsize=50
			if rst.fields(i).Name="����" then 
				textvalue="ҩƷ"
			else	
				textvalue=""
			end if	
		elseif rst.Fields(i).type=3 then
			texttype="������"
			textmaxlength=10
			textsize=10
			textvalue=0
		elseif rst.Fields(i).Type=7 then
			texttype="����/ʱ��"
			textmaxlength=20
			textsize=20
			textvalue=now()
		end if
		msg=msg&"<tr><td>"&rst.Fields(i).Name&"("&texttype&")</td><td><input type=text name='text"&i&"' value="&chr(34)&textvalue&chr(34)&" size='"&textsize&"' maxlength='"&textmaxlength&"'></td></tr>"
	next	
else
	if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
	for i=1 to rst.Fields.Count-1	
		if rst.Fields(i).type=202 then
			texttype="�ַ���"
			textmaxlength=rst.Fields(i).DefinedSize
			textsize=textmaxlength
			if textsize>50 then textsize=50
			textvalue=rst.Fields(i).Value
		elseif rst.Fields(i).type=3 then
			texttype="������"
			textmaxlength=10
			textsize=10
			textvalue=rst.Fields(i).Value
		elseif rst.Fields(i).Type=7 then
			texttype="����/ʱ��"
			textmaxlength=20
			textsize=20
			textvalue=rst.Fields(i).Value
		end if
		msg=msg&"<tr><td>"&rst.Fields(i).Name&"("&texttype&")</td><td><input type=text name='text"&i&"' value="&chr(34)&textvalue&chr(34)&" size='"&textsize&"' maxlength='"&textmaxlength&"'></td></tr>"
	next	
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"<tr align=center><td colspan=2><input type=submit name=submit value='"&opt&"'> <input type=button onclick='javascript:history.back();' value='����'></td></tr></table></body>"
Response.Write msg
%>