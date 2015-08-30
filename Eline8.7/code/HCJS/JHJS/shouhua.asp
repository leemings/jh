<%@ LANGUAGE=VBScript codepage ="936" %>

<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
if Session("sjjh_inthechat")<>"1" then %>
<script language="vbscript">
MsgBox "请进入聊天室再进行出售鲜花操作！！"
window.close()
</script>
<%response.end
end if%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>鲜花收购处♀wWw.happyjh.com♀</title>
<link rel="stylesheet" href="../../css.css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="7">
<div align="center">
<table width="778" border="0" cellspacing="0" cellpadding="4">
<tr>
<td width="101" align="center"><img src="shouhua.gif" width="101" height="304"></td>
<td colspan="2" valign="top">
<p align="center">[ <font color="#FF3300">鲜 花 收 购</font> ]</p> 
 
<% 
Set conn=Server.CreateObject("ADODB.CONNECTION") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
conn.open Application("sjjh_usermdb") 
rs.open "SELECT w7 FROM 用户 WHERE 姓名='" & sjjh_name& "'",conn 
 
if IsNull(rs("w7"))  then 
	rs.close 
	set rs=nothing 
	conn.close 
	set conn=nothing 
	 
 
%> 
	老板惊叫道： 
<table width="450" align="center" cellspacing="2" border="1" 
cellpadding="5" bordercolor="#000000"> 
<tr> 
<td><font color="#FF6600">天哪</font>——:-O<br> 
<%=name%>你没出息，不种花，一朵花也没有，来这里做什么？勤劳点吧！！</td> 
</tr> 
</table> 
<% 
else 
%><%=name%> 您目前拥有的鲜花列表： 
<table border="0" cellspacing="0" cellpadding="0" width="90%" align="center"> 
<tr > 
<td style="border-bottom: 1px solid rgb(250,50,50)"> 
	<div align="center"><font color="#FF6600">鲜花名称</font></div> 
</td> 
<td style="border-bottom: 1px solid rgb(250,50,50)"> 
	<div align="center"><font color="#FF6600">类  型</font></div> 
</td> 
<td style="border-bottom: 1px solid rgb(250,50,50)"> 
	<div align="right"><font color="#FF6600">数  量</font></div> 
</td> 
<td style="border-bottom: 1px solid rgb(250,50,50)"> 
	<div align="right"><font color="#FF6600">收购价格</font></div> 
</td> 
<td style="border-bottom: 1px solid rgb(250,50,50)"> 
	<div align="center"><font color="#FF6600">操  作</font></div> 
</td> 
</tr> 
<% 
d7=rs("w7") 
 
rs.close 
set rs=nothing 
conn.close 
set conn=nothing 
 
 
if not IsNull(d7) then 
	call show(d7,"鲜花") 
end if 
 
 
end if 
 
sub show(data,lx) 
Set conn=Server.CreateObject("ADODB.CONNECTION") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
conn.open Application("sjjh_usermdb") 
data1=split(data,";") 
data2=UBound(data1) 
 
for y=0 to data2-1 
	data3=split(data1(y),"|") 
	rs.open "SELECT h FROM b WHERE a='" & data3(0)& "'",conn 
	if  not rs.bof and not rs.eof  then 
%> 
  <tr bgcolor="#dfefff" onmouseout="this.bgColor='#DFEFFF';"onmouseover="this.bgColor='#85C2E0';">  
    <td>  
      <div align="center"><%=data3(0)%> </div> 
    </td> 
	 <td>  
      <div align="center"><%=data3(0)%> </div> 
    </td> 
    <td>  
      <div align="right"><%=data3(1)%> </div> 
    </td> 
    <td> 
      <div align="right"><%=rs("h")%><%if data3(0)="冰雪" or data3(0)="逸韵" or data3(0)="香芸" or data3(0)="忧梦" or data3(0)="烈焰" then%>元<%else%>两<%end if%>/个</div> 
    </td> 
    <td> 
      <div align="right"><%if data3(0)="冰雪" or data3(0)="逸韵" or data3(0)="香芸" or data3(0)="忧梦" or data3(0)="烈焰" then%><%=int(rs("h")/4)%>元<%else%><%=int(rs("h")/3)%>两<%end if%>/个 
      </div> 
    </td> 
    <td>  
	    <div align="center"> <a href="dan2.asp?wpname=<%=data3(0)%>&lx=<%=data3(0)%>" ><font color="#0000FF">卖出</font></a> </div> 
    </td> 
  </tr> 
  <% 
  end if 
  rs.close 
erase data3 
next 
erase data1 
set rs=nothing 
conn.close 
set conn=nothing 
end sub 
%> 
 
</table> 
</td> 
</tr> 
</table> 
<br>版权：『快乐江湖』&#8482;　版本：ELINE 7.0.0　站长：回首当年</div>
</body>
</html>