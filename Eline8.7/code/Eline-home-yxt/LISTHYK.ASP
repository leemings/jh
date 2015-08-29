<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
lx=request("lx")
%>
<html>
<head>
<title>会员数据库管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center"><font color="#0000FF">会员卡管理</font></p>
<p align="center"><font color="#0000FF">&nbsp;</font></p>
<p align="center">　</p>
<table border="1" width="73%" bordercolorlight="#000000" cellspacing="1" cellpadding="1" bordercolordark="#85C2E0" height="25" align="center">
  <tr> 
    <td  width="18%" align="center" nowrap height="25">
    卡号</td>
    <td  width="16%" align="center" nowrap height="25">
    剩余点数</td>
    <td  width="16%" align="center" nowrap height="25">
    当前状态</td>
    <td  width="19%" align="center" nowrap height="25">
    所有者ID</td>
    <td  width="15%" align="center" nowrap height="25">
    用户剩余时间</td>
    <td  width="17%" align="center" nowrap height="25">
    操作</td>
  </tr>
<%

if lx=null then
%>
 <tr><td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="101%" align="center" nowrap height="25" colspan="6"><font color="ff0000">您的操作是非法的!请不要在尝试了!</font></td></tr>

<%
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_hykmdb")
select case lx
case 1
rs.open "SELECT * FROM 会员卡 Where 类型=0",conn
case 2
rs.open "SELECT * FROM 会员卡 Where 类型<>5",conn
case 3
hykcx=Request.form("hykcx")
if hykcx="卡号" THEN 
rs.open "SELECT * FROM 会员卡 Where 卡号='"&Request.form("hykcxname")&"'",conn
else
rs.open "SELECT * FROM 会员卡 Where 用户ID="&INT(ABS(clng(Request.form("hykcxname"))))&"",conn
END IF
end select
if rs.eof or rs.bof then
%>
 <tr><td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="101%" align="center" nowrap height="25" colspan="6"><font color="ff0000">抱歉,没有找到符合你要求的记录!</font></td></tr>
<%
else
do while not rs.bof and not rs.eof
%>
 <tr bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';" > 
    <td  width="18%" align="center" nowrap height="25"><%=rs("卡号")%></td>
    <td  width="16%" align="center" nowrap height="25"><%=rs("点数")%></td>
    <td  width="16%" align="center" nowrap height="25">
    <%
    select case rs("类型")
    case 0
    %>
    未使用 
    <% 
   case 1 
    %> 
    已使用 
    <% 
   case 2 
    %> 
   冲泡点分钟 
    <% 
   case 3 
    %>  
   冲等级分钟 
    <% 
   case 4 
    %>  
   已作废 
    <% 
end select 
    %> 
    </td> 
    <td  width="19%" align="center" nowrap height="25"><%=rs("用户ID")%></td> 
    <td  width="15%" align="center" nowrap height="25"><%=rs("时间")%></td> 
    <td  width="17%" align="center" nowrap height="25"> 
    <a href="showhyk.asp?ID=<%=rs("卡号")%>">修改</a> </td> 
  </tr> 
<% 
rs.movenext 
loop  
end if  
rs.close 
conn.close 
set rs=nothing 
set conn=nothing 
end if 
%> 
</table> 
<p align="center">　</p> 
<p align="center">本程序由阿男编写</p>
</body>
</html>