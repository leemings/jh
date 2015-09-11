<%@ LANGUAGE=VBScript codepage ="936" %>
<html><title>npc管理</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666 oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
npcid=request.querystring("npcid")
npcid=cint(npcid)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sqlstr="SELECT * FROM npc where id="&npcid
rs.Open sqlstr,conn
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('抱歉你所要查找的npc我们找不到！请查看是否正确！');history.back();</script>"
end if
%>
<form method=post action=npc_update.asp>
<input type=hidden size=4 name=id value="<%=npcid%>">
<br>修改说明：<br>1、[n图片]：所在目录 chat 文件夹。格式：“ npc/gui2.gif ”（引号内的）。<br>2、[n物品]：打死“npc”后，所随机得到的物品。格式必须按照：例如“ 白饭|1;九花雨露丸|1;桃酥|1; ”（引号内的）。<br>修改添加，注意“;”为半角，字符之间无空格。<br>3、添加修改“爆出物品”时，该物品，必须是江湖中有的；否则，“爆出物品”无法使用。<br>4、[n物品类型]：既爆出物品的类型。<br>w1为药品类，w2为毒药类，w4为暗器类，w5为卡片类，w6为物品类，w7为鲜花类，w8为配药类，w9为魔宝类，w10法器为类。<br>z1为头盔，z2为装饰，z3为盔甲，z4为左手，z5为右手和锻造，z6为双脚。<br>5、爆出物品与物品类型，必须是相对应的；否则，打死npc后，得到的“物品”在“个人物品”中就无法显示。<br>6、[n出场词]：图片“&lt;img src=npc/jiu2.gif width=50&gt;”所在目录 chat 文件夹。格式：“ &lt;img src=npc/jiu2.gif width=50&gt; ”（引号内的）。<br><br><table border="1" width="100%" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center"><tr><td colspan='6'><div align='center'>怪物ID:<%=npcid%>(<font color=blue size=-1>注意：数据类型与所对应的数据必需相同，否则出错！不能有空的数据存在!----原来就没有的可以打空格!----</font>)</div></td></tr>
<%
Response.Write "<tr>"
for i=1 to rs.Fields.Count-1
strtype=rs.fields(i).Type
strname=rs.Fields(i).Name
if strname="姓名" then seename=rs.Fields(i).value
if rs.fields(i).Type=202 then strtype="字符"
if rs.fields(i).Type=3 then strtype="数字"
if rs.fields(i).Type=7 then strtype="日期"
if rs.fields(i).Type=11 then strtype="逻辑"
if strname="grade" then strname="管理等级"
if strname="mvalue" then strname="月积分"
if strname="allvalue" then strname="总积分"
if strname="times" then strname="登录次数"
if strname="regip" then strname="注册ip"
if strname="lastip" then strname="最后ip"
if strname="regtime" then strname="注册时间"
if strname="lasttime" then strname="最后时间"
	Response.Write "<td><div align='right'><font size=-1>"&strname&"(<font color=blue >"&strtype&"</font>)：<input type=text size='20' style='FONT-SIZE: 9pt;' name="&rs.Fields(i).Name&" value='"&rs.Fields(i).value&"'></font></div></td>"
	if i/3=int(i/3) then Response.Write "</tr><tr>"
	next
Response.Write "</tr>"
Response.Write "<tr><td colspan='6'><div align='center'><input type=submit value=更新 name=submit> <input type=submit value=新增 name=submit> <input type=submit value=删除 name=submit> <input type=reset value=重置></div></td></tr>"
Response.Write "</table></form>"
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>