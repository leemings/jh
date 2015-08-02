<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=Session("Ba_jxqy_usercorp")
if username="" then Response.redirect "../error.asp?id=016"
attackname=Server.HTMLEncode (Trim(Request.Form("attackname")))
especial=Server.HTMLEncode (Trim(Request.Form("especial")))
energy=Trim(Request.Form("energy"))
if instr(attackname,"'")<>0 or instr(attackname," ")<>0 or instr(attackname,"\")<>0 or instr(attackname,"/")<>0 or instr(attackname,chr(34))<>0 then Response.Redirect "../error.asp?id=056"
if instr(";无;麻痹;中毒;疯狂;封咒;",";"&especial&";")=0 then Response.Redirect "../error.asp?id=024"
if not isnumeric(energy) then Response.Redirect "../error.asp?id=024"
energy=clng(energy)
if energy<10 then energy=10
if especial="无" then 
	mp=energy\10
else
	mp=energy
end if		
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 用户 where 姓名='"&username&"' and 门派='"&usercorp&"' and 身份='掌门' and 精力>="&energy*20,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=065"
rst.Close
on error resume next
conn.BeginTrans
conn.Execute "insert into 招式(招式,门派,修习条件,消耗精力,消耗内力,基本攻击,特效,说明,攻击说明) values('"&attackname&"','"&usercorp&"','True',"&energy&","&mp&","&energy&",'"&especial&"','"&username&"所创，适合所有人修习。','##对%%使用了"&usercorp&"的独门武功"&attackname&"')"
conn.Execute "update 用户 set 精力=精力-"&energy*20&" where 姓名='"&username&"'"
if conn.Errors.Count>0 then
	conn.RollbackTrans
	Response.Redirect "../error.asp?id=104&errormsg="&conn.Errors(0).Description
else
	conn.CommitTrans
end if		
conn.Close
set conn=nothing
%>
<script language=javascript>history.back();</script>