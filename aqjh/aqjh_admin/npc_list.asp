<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html><title>npc列表</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666 oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false onselectstart=event.returnValue=false>
</head>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id,n姓名,n性别,n图片,n经验,n攻击,n防御,n武功,n体力,n内力,n物品,n银两,n物品类型,n物品,n主人 from [npc] order by n经验",conn,1,1
if rs.eof and rs.bof then
	response.write "<p align='center'>没有可排行的对象 </p>"
else
npcsl=rs.recordcount
%>
<table border="1" width="90%" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
  <tr bgcolor="#336633"> 
    <td colspan="13" align=center><font color="#FFFFFF">NPC 列 表</td>
  </tr>
  <tr bgcolor="#336633"> 
    <td width="5%" align=center><font color="#FFFFFF">ID号</td>
    <td width="10%" align=center><font color="#FFFFFF">姓名</td>
    <td width="4%" align=center><font color="#FFFFFF">性别</td>
    <td width="5%" align=center><font color="#FFFFFF">经验</td>
    <td width="6%" align=center><font color="#FFFFFF">攻击</td>
    <td width="4%" align=center><font color="#FFFFFF">防御</td>
	<td width="4%" align=center><font color="#FFFFFF">武功</td>
    <td width="9%" align=center><font color="#FFFFFF">体力</td>
    <td width="7%" align=center><font color="#FFFFFF">内力</td>
    <td width="9%" align=center><font color="#FFFFFF">类型</td>
    <td width="5%" align=center><font color="#FFFFFF">操作</td>
  </tr>
<%do while not rs.eof%>
  <tr> 
    <td width="5%" align=center><%=rs("id")%></td>
    <td width="10%" align=center><%=rs("n姓名")%></td>
    <td width="4%" align=center><%=rs("n性别")%></td>
    <td width="5%" align=center><%=rs("n经验")%></td>
    <td width="6%" align=right><%=rs("n攻击")%></td>
    <td width="4%" align=right><%=rs("n防御")%></td>
    <td width="4%" align=right><%=rs("n武功")%></td>
    <td width="9%" align=right><%=rs("n体力")%></td>
    <td width="7%" align=right><%=rs("n内力")%></td>
    <td width="9%" align=left><%=rs("n物品类型")%></td>
    <td width="5%" align=center><a href="npc_show.asp?npcid=<%=rs("id")%>">详细</a></td>
  </tr>
<%
rs.movenext
loop
%>
</table>
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>