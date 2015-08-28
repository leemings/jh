<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
cur=Request.QueryString("cur")
if instr(cur,"'")<>0 then Response.End
if username="" then
msg="top.location.location.href='../error.asp?id=016';"
else
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 物品 where 所有者='"&username&"' and 属性='药品' and 数量>0 and 名称='"&cur&"'",conn
if rst.EOF or rst.BOF then
msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>【用药】</FONT>您没有"&cur&"<br>');"
else
hpadd=rst("体力")
mpadd=rst("内力")
rst.Close
rst.Open "select 体力,内力,等级,攻击,防御 from 用户 where 姓名='"&username&"'",conn
dj=rst("等级")
tili=rst("体力")
neili=rst("内力")
if hpadd>0 and tili>dj*1000000 then
msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>【用药】</FONT>现在您的体力上限应该是"&dj*1000000&",无法再吃加体力的药了,小心撑死<br>');"
elseif mpadd>0 and neili>dj*100000 then
msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>【用药】</FONT>现在您的内力上限应该是"&dj*100000&",无法再吃加内力的药了,小心撑死<br>');"
else
conn.Execute "update 用户 set 体力=体力+"&hpadd&",内力=内力+"&mpadd&" where 姓名='"&username&"'"
conn.Execute "update 物品 set 数量=数量-1 where 所有者='"&username&"' and 名称='"&cur&"'"
msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>【用药】</FONT>您使用了"&cur&",体力增加"&hpadd&",内力增加"&mpadd&"<br>');parent.confrm.document.form1.hp.value='生命："&hp+hpadd&"';parent.confrm.document.form1.mp.value='内力："&mp+mpadd&"';parent.optfrm.location.replace('curative.asp');"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
end if
end if
Response.Write "<script language=javascript>"&msg&"</script>"
%>
