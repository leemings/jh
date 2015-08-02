<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.redirect "../error.asp?id=016"
corpcond=server.HTMLEncode(Trim(Request.Form("corpcond")))
if instr(corpcond,"'")<>0 or instr(corpcond,chr(34))<>0 then Response.Redirect "../error.asp?id=056"
tj=Request.Form("corptj")
corpname=server.HTMLEncode(Trim(Request.Form("corpname")))
if corpname="" or corpname="ÎŞ" or instr(corpname,"'")<>0 or instr(corpname,chr(34))<>0 or instr(corpname," ")<>0 or instr(corpname,"ÿ")<>0 or instr(corpname,"&#63733;") then Response.Redirect "../error.asp?id=056"
corpsilver=server.HTMLEncode(Trim(Request.Form("corpsilver")))
if not isnumeric(corpsilver) then Response.Redirect "../error.asp?id=024"
silver=clng(corpsilver)
if silver<0 then 
	silver=0
elseif silver>1000 then
	silver=1000
end if
if tj="" then Response.Redirect "../error.asp?id=102"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("Adodb.recordset")
rst.Open "select * from ÓÃ»§ where ĞÕÃû='"&username&"' and ÃÅÅÉ<>'¹Ù¸®' and Éí·İ<>'ÕÆÃÅ' and ¹¥»÷>100000 and ·ÀÓù>100000 and ÌåÁ¦>100000 and ÄÚÁ¦>100000 and ÒøÁ½>"&silver*100000&" and ¾«Á¦>10000",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=063"
rst.Close
rst.Open "select * from ÃÅÅÉ where ÃÅÅÉ='"&corpname&"'",conn,1,2
if rst.EOF or rst.BOF then
	rst.AddNew
	rst("ÃÅÅÉ")=corpname
	rst("¹¤×ÊÏµÊı")=silver
	rst("¼ò½é")=corpcond
	rst("¼ÓÈëÌõ¼ş")=tj
	rst("chaton")=false
	rst.Update
	conn.Execute "update ÓÃ»§ set ¾«Á¦=¾«Á¦-10000,ÒøÁ½=ÒøÁ½-"&silver*100000&",Éí·İ='ÕÆÃÅ',ÃÅÅÉ='"&corpname&"' where ĞÕÃû='"&username&"'"
	Session("Ba_jxqy_usercorp")=corpname
else 
	Response.Redirect "../error.asp?id=064"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing 
%>
<head><link rel='stylesheet' href='../chatroom/style1.css'></head><body oncontextmenu=self.event.returnValue=false background='<%=bgimage%>' bgcolor='<%=bgcolor%>' topmargin=100><p align=center><font color=ff0000 size=5>×ÔÁ¢ÃÅ»§</font></p><hr>
¹§Ï²¹§Ï²£¬¾­¹ı³¤ÆÚµÄ·Ü¶·,ÄãÖÕÓÚ½¨Á¢ÁË×Ô¼ºµÄÃÅÅÉ<font color=0000ff size=6><%=corpname%></font> ¾«Á¦¼õÉÙ10000,ÒøÁ½¼õÉÙ<%=silver*100000%><br><p align=center><a href=javascript:location.replace('addcorp.asp');>·µ»Ø</a></p>
