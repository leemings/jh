<!--#include file="data.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql1="Select  * from 用户 where 姓名='"&username&"'"  
set rs1=conn.Execute(sql1)
id=rs1("id")
sex=rs1("性别")
money=rs1("银两")
peiou=rs1("配偶")
if peiou<>"无" then Response.Redirect "../error.asp?id=039"
if money<5000 then Response.Redirect "../error.asp?id=030"
if sex="男" then 
sex1="女"
else
sex1="男"
end if
sql="Select  * from 用户 where 性别='"&sex1&"'" 
set rs=conn.Execute(sql)
if rs.eof and rs.bof then Response.Redirect "../error.asp?id=515"
total=rs.RecordCount
randomize timer
rand=100+Int(id * Rnd)
if rand<1 then rand=001
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="Select  * from 用户 where 性别='"&sex1&"' and id="&rand&"" 
set rs=conn.Execute(sql)
if rs.eof and rs.bof then Response.Redirect "../error.asp?id=515"
%>
<html>

<head>
<title>速配成功</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">
<!--
--></style>
<link rel="stylesheet" href="../style4.css">
</head>

<body bgcolor="#3a4b91" topmargin="0" leftmargin="0" text="#FFFFFF">
<div align="center"><font color="#FFFF00"><br>
  </font><font color="#00FF00">
  你速配成功了，我们已经帮你向对方刊登了求婚信息，并收取了你5000大两！</font><br>
</div>
<br>
<table border="1" cellspacing="1" cellpadding="0" align="center"
width="429" bordercolor="#000000">
<tr >
<td colspan="2">
<div align="center"><font
size="2" color="#FFFFFF"> 你的速配对象<b><font color="#FFFFFF"><%=rs("姓名")%></font></b>
<% if rs("性别")="男" then%>
<font size="2" color="#FFFFFF">大侠</font></font>
<% else %> 
<font size="2" color="#FFFFFF">(侠女)<font size="3" ></font></font> 
<% end if %> 
<font size="2" color="#FFFFFF" class="c">的剑侠背景</font> </div> 
</td> 
</tr> 
<tr> 
<td width="64"><font class="c" color="#FFFFFF">姓 名</font></td> 
<td width="219"><font color="#FFFFFF"><%=rs("姓名")%></font></td> 
</tr> 
<tr> 
<td width="64"><font class="c" color="#FFFFFF">性 别</font></td> 
<td width="219"><font color="#FFFFFF"><%=rs("性别")%> </font> </td> 
</tr> 
<tr> 
<td width="64"><font class="c" color="#FFFFFF">Email地址</font></td> 
<td width="219"><font color="#FFFFFF"><%=rs("电子邮箱")%> </font> </td> 
</tr> 
</table> 
<table border="1" cellspacing="1" cellpadding="0" align="center" 
width="429" bordercolor="#000000"> 
<tr> 
<td colspan="4" height="20"> 
<div align="center"> <font size="2" color="#FFFFFF">剑 侠 档 案</font> </div> 
</td> 
</tr> 
<tr> 
<td width="64" height="25"><font color="#FFFFFF">现 金：</font></td>  
<td width="141" height="25"><font color="#FFFFFF"><%=rs("银两")%> 两</font></td> 
<td width="66" height="25"><font color="#FFFFFF">精 力：</font></td>  
<td width="127" height="25"><font color="#FFFFFF"><%=rs("精力")%></font></td> 
</tr> 
<tr> 
<td width="64" height="20"><font color="#FFFFFF">存 款：</font></td>  
<td width="141" height="24"><font color="#FFFFFF">保密</font></td> 
<td width="66" height="24"><font color="#FFFFFF">道 德：</font></td>  
<td width="127" height="24"><font color="#FFFFFF"><%=rs("道德")%></font></td> 
</tr> 
<tr> 
<td width="64" height="20"><font color="#FFFFFF">积 分：</font></td>  
<td width="141" height="24"><font color="#FFFFFF"><%=rs("积分")%></font></td> 
<td width="66" height="24"><font color="#FFFFFF">攻 击：</font></td>  
<td width="127" height="24"><font color="#FFFFFF"><%=rs("攻击")%></font></td> 
</tr> 
<tr> 
<td width="64" height="20"><font color="#FFFFFF">内 力：</font></td>  
<td width="141"><font color="#FFFFFF"><%=rs("内力")%></font></td> 
<td width="66"><font color="#FFFFFF">防 御：</font></td>  
<td width="127"><font color="#FFFFFF"><%=rs("防御")%></font></td> 
</tr> 
<tr> 
<td width="64" height="20"><font color="#FFFFFF">资 质：</font></td>  
<td width="141"><font color="#FFFFFF"><%=rs("资质")%></font></td> 
<td width="66"><font color="#FFFFFF">门 派：</font></td>  
<td width="127"><font color="#FFFFFF"><%=rs("门派")%> </font> </td> 
</tr> 
<tr> 
<td width="64" height="20"><font color="#FFFFFF">体 力：</font></td>  
<td width="141"><font color="#FFFFFF"><%=rs("体力")%></font></td> 
<td width="66"><font color="#FFFFFF">身 份：</font></td>  
<td width="127"><font color="#FFFFFF"><%=rs("身份")%></font></td> 
</tr> 
<tr> 
<td width="64" height="20"><font color="#FFFFFF">配 偶：</font></td>  
<td width="141"><font color="#FFFFFF"><%=rs("配偶")%></font></td> 
<td width="66"><font color="#FFFFFF">等 级：</font></td>  
<td width="127"><font color="#FFFFFF"><%=rs("等级")%> </font> </td> 
</tr> 
<tr> 
<td width="64" height="20"><font color="#FFFFFF">状态：</font></td> 
<td width="141"><font color="#FFFFFF"><%=rs("状态")%> 
  </font> 
</td> 
<td width="66"><font color="#FFFFFF">荣誉：</font></td>
<td width="127"><font color="#FFFFFF"><%if rs("等级")="" then response.write("武林新手")                                                          
if rs("等级")<=3 then response.write("武林新手")                                                          
if rs("等级")<=6 and rs("等级")>3  then response.write("武林游侠")                                                          
if rs("等级")<=10 and rs("等级")>6  then response.write("武林斗侠")                                                          
if rs("等级")<=15and rs("等级")>10  then response.write("初级战将")                                                          
if rs("等级")<=20 and rs("等级")>15  then response.write("中级战将")                                                          
if rs("等级")<=25 and rs("等级")>20  then response.write("高级战将")                                                          
if rs("等级")<=30 and rs("等级")>25  then response.write("超级战将")                         
if rs("等级")<=35 and rs("等级")>30  then response.write("武神")                                                            
if rs("等级")<38 and rs("等级")>35 then response.write("战神")                          
if rs("等级")>=38 then response.write("终极战神")                                                          
%></font></td>
</tr>
</table><div align="center"><font color="#FFFFFF"><br> 
  </font> 
</div> 
</body> 
 
 
</html> 
<% 
name2=rs("姓名") 
mess="本求婚信息是经速配成功后速配中心代为刊登。" 
sql="update 用户 set 银两=银两-5000 where 姓名='" &username& "'" 
conn.execute sql 
sz = "'" & username & "','" & name2 & "','" & mess & "', now()" 
                		into_db = "INSERT INTO 媒婆 (申请人, 对象, 说明, 时间) VALUES(" & sz & ")" 
	                	connt.Execute(into_db) 
rs.Close             
set rs=nothing             
conn.Close             
set conn=nothing
set connt=nothing           
%> 
