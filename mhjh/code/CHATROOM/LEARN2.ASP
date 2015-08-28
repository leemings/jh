<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("yx8_mhjh_username")
if un="" then Response.Redirect "../error.asp?id=016"
mg=Request.QueryString("mg")
if instr(mg,"'")<>0 then Response.Redirect "../error.asp?id=024"
co=Session("yx8_mhjh_usercorp")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 会员 from 用户 where 姓名='"&un&"'",conn
hy=rst("会员")
rst.Close
rst.Open "select 等级,消耗精力,特效,基本攻击 from 攻击 where 姓名='"&un&"' and 招式='"&mg&"'",conn
if rst.EOF or rst.BOF then
rst.Close
rst.Open "select * from 招式 where 门派='"&co&"' and 招式='"&mg&"'",conn
if rst.EOF or rst.BOF then
msg="你想修习"&mg&"，而本门没有此样招式可供修习。"
else
energy=rst("消耗精力")
proviso=rst("修习条件")
basemp=rst("消耗内力")
baseat=rst("基本攻击")
especial=rst("特效")
atdeclaration=rst("攻击说明")
rst.Close
rst.Open "select * from 用户 where 姓名='"&un&"' and 精力>="&energy&" and "&proviso,conn
if rst.EOF or rst.BOF then
provisotxt=replace(proviso,"and","并且")
provisotxt=replace(provisotxt,"or","或者")
provisotxt=replace(provisotxt,">=","不低于")
provisotxt=replace(provisotxt,"<=","不高于")
provisotxt=replace(provisotxt,">","大于")
provisotxt=replace(provisotxt,"<","小于")
provisotxt=replace(provisotxt,"=","为")
msg="你修习"&mg&"时，虽有高人指点，苦于自己资质有限。因而失败，需精力"&energy&"，条件满足："&provisotxt
else
conn.Execute "insert into 攻击(姓名,招式,等级,消耗精力,消耗内力,基本攻击,特效,攻击说明) values('"&un&"','"&mg&"',1,"&energy&","&basemp&","&baseat&",'"&especial&"','"&atdeclaration&"')"
conn.execute "update 用户 set 精力=精力-"&energy&" where 姓名='"&un&"'"
msg="你修习"&mg&"第1级成功。耗用精力"&energy
end if
end if
elseif rst("基本攻击")>1900000000 then Response.write "<script language='javascript'>alert ('你的这招武功的威力已经大于19亿了，无法继续修炼');location.href='learn.asp';</script>"
elseif (rst("等级")<120 and hy=True) or (rst("等级")<120 and co="官府") then
energy=rst("消耗精力")
agrade=rst("等级")
rst.Close
rst.Open "select * from 用户 where 姓名='"&un&"' and 精力>="&energy*agrade,conn
if rst.EOF or rst.BOF then
msg="你修习"&mg&"第"&agrade&"级时心有旁骛失败，无法专心入定，因而失败，需精力"&energy
else
conn.Execute "update 用户 set 精力=精力-"&energy*agrade*0.5&" where 姓名='"&un&"'"
conn.Execute "update 攻击 set 等级=等级+1,基本攻击=基本攻击*1.1 where 姓名='"&un&"' and 招式='"&mg&"'"
msg="你修习"&mg&"第"&agrade+1&"级成功。耗用精力"&energy*agrade*0.5
end if
elseif rst("等级")<100 then
energy=rst("消耗精力")
agrade=rst("等级")
rst.Close
rst.Open "select * from 用户 where 姓名='"&un&"' and 精力>="&energy*agrade,conn
if rst.EOF or rst.BOF then
msg="你修习"&mg&"第"&agrade&"级时心有旁骛失败，无法专心入定，因而失败，需精力"&energy
else
conn.Execute "update 用户 set 精力=精力-"&energy*agrade&" where 姓名='"&un&"'"
conn.Execute "update 攻击 set 等级=等级+1,基本攻击=基本攻击*1.1 where 姓名='"&un&"' and 招式='"&mg&"'"
msg="你修习"&mg&"第"&agrade+1&"级成功。耗用精力"&energy*agrade
end if
else	
msg="你已经修习"&mg&"至最高境界了，无法再行修练。"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=replace(msg,"\","\\")
msg=replace(msg,"/","\/")
msg=replace(msg,chr(34),"\"&chr(34))
msg=replace(msg,chr(39),"\"&chr(39))
%>
<html>
<head>
<title>修习武功</title>
<link rel=stylesheet href='css2.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='bg1.gif' >
<div align=center>
<font color=0000ff size=5>修习武功</font>
<hr>
<input type=button value='返回' onclick="javascript:location.href='learn.asp'">
</div>
<%=msg%>
</body>
</html>
