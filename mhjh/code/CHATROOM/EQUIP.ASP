<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
commodityname=Request.QueryString("commodityname")
if instr(commodityname,"'")<>0 or instr(commodityname," ")<>0 then Response.End
gj=session("yx8_mhjh_usergrade")
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "chaterror.asp?id=000"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select tc.ID,tc.属性,tc.攻击,tc.防御,tc.特效,tu.特技 from 物品 tc inner join 用户 tu on tu.姓名=tc.所有者 where tc.名称='"&commodityname&"' and tc.所有者='"&username&"' and tc.数量>0"
rst.Open sqlstr,conn
cmid=rst("id")
cmtype=rst("属性")
cmattack=rst("攻击")
cmdefence=rst("防御")
cmespecial=rst("特效")
especial=rst("特技")
if rst.EOF or rst.BOF then
msg="你没有此样物品可供装备"
end if
if (commodityname="天王之刃" or commodityname="天王之铠") and dj<20 then
msg="你等级不够20级,无法装备此物品"
elseif (commodityname="电光之剑" or commodityname="电光之铠") and dj<40 then
msg="你等级不够40级,无法装备此物品"
elseif (commodityname="破坏之剑" or commodityname="破坏之铠") and dj<60 then
msg="你等级不够60级,无法装备此物品"
elseif (commodityname="魔幻之剑" or commodityname="魔幻之铠") and dj<80 then
msg="你等级不够80级,无法装备此物品"
else
rst.Close
rst.Open "select id from 物品 where 属性='"&cmtype&"' and 装备=true and 所有者='"&username&"'",conn
if rst.EOF or rst.BOF then
if cmespecial="无" or isnull(cmespecial) or cmespecial="" then
especial=especial
else
especial=especial&cmespecial&";"
end if
conn.Execute "update 用户 set 攻击=攻击+"&cmattack&",防御=防御+"&cmdefence&",特技='"&especial&"' where 姓名='"&username&"'"
conn.Execute "update 物品 set 数量=数量-1,装备=True where id="&cmid
msg="你成功的装备了"&commodityname&",攻击因些而增加"&cmattack&",防御因此而增加"&cmdefence
else
msg="错误，你已经装备有"&cmtype&",无法再行装备"&commodityname
end if
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><title>装备物品</title><link rel="stylesheet" href="css.css"><script language=javascript>setTimeout("location.replace('seeequip.asp');",3000);</script></head>
<body  oncontextmenu=self.event.returnValue=false bgcolor="#FFDDF2" background='bg1.gif'>
<p align=center><font color=0000FF>装备物品</font></p>
<p align="center">
<font color=FF0000><%=msg%></font>
<br><br><br>3秒钟后自动返回<br><a href="javascript:location.replace('seeequip.asp');" onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></body>
