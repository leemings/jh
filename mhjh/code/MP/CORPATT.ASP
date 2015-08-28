<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
usercorp=Session("yx8_mhjh_usercorp")
if username="" then Response.redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 精力 from 用户 where 姓名='"&username&"' and 门派='"&usercorp&"' and 身份='掌门'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=046"
userenergy=rst("精力")
rst.Close
rst.Open "select * from 招式 where 门派='"&usercorp&"' order by 基本攻击",conn
do while not rst.EOF
msg=msg&"<tr><td>"&rst("招式")&"</td><td>"&rst("消耗精力")&"</td><td>"&rst("消耗内力")&"</td><td>"&rst("基本攻击")&"</td><td>"&rst("特效")&"</td></tr>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel='stylesheet' href='../style.css'>
<script language=javascript>var userenergy=<%=userenergy\20%>;function energychg(){if(document.form1.energy.value>userenergy){alert('您的精力不足！');}else{if(document.form1.especial.value=='无'){document.form1.mp.value=parseInt(document.form1.energy.value/10);}else{document.form1.mp.value=document.form1.energy.value;}document.form1.attack.value=document.form1.energy.value;}}</script>
<title>武功设置</title>
</head>
<body oncontextmenu=self.event.returnValue=false  background='../chatroom/bg1.gif'><p align=center><font color="ff0000" size="4">武功设置</font><form action=corpatt2.asp method=post id=form1 name=form1><table width=80% align=center border=1 cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF"><tr align=center bgcolor=ffffff><td>招式</td><td>精力</td><td>内力</td><td>攻击</td><td>特效</td></tr><%=msg%>
<tr><td><input type=text name=attackname value='填写招式名称' maxlength=7 size=14 ></td><td><input type=text name=energy value='<%=userenergy\20%>' maxlength=5 size=5  onchange='javascript:energychg();'></td><td><input type=text name=mp value='<%=userenergy\200%>' maxlength=5 size=5 readonly></td><td><input type=text name=attack value='<%=userenergy\20%>' maxlength=5 size=5 readonly></td><td><select name=especial onchange='javascript:energychg();'><option value='无' selected>无</option><option value='麻痹'>麻痹</option><option value='中毒'>中毒</option><option value='封咒'>封咒</option><option value='疯狂'>疯狂</option></select></td></tr></table><p align=center><input type=submit value='新建' id=submit1 name=submit1></p></form></body>
