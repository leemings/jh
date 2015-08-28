<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
timetmp=now()
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
name=Session("yx8_mhjh_userName")
sql="select 任务时间,任务,风铃 from 用户 where 姓名='" & name & "'"
set rs=conn.execute(sql)
if rs("风铃")<1 then
%>
<script language=vbscript>
MsgBox "我倒，你连一个风铃都没有呀,想办法弄一个再来吧！"
location.href = "javascript:self.close()"
</script>
<%
else
if rs("风铃")>0 then
sql="update 用户 set 风铃=风铃-1,任务时间='" & timetmp & "',任务='孤岛探宝' where 姓名='" & name & "'"
conn.execute sql
dim newtalkarr(600) 
   talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username 
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>【江湖通报】</font><font color=blue>"&name&"使用了一个风铃,成功进入海外孤岛,行动开始时间是" & timetmp & "让我们祝他好运吧!</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
end if
end if
%>
<html>
<head>
<title>魔幻宝物调查所</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../../style.css">
</head>

<body  background='../../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5">
<div align="center"><font color="#FFFFFF"><br>
</font>
<font size="2"> [你一个风铃来到了魔幻海岛宝物调查所，这里有各种江湖人士从不同的地方打探到的藏宝消息]<br>
</font><font size="2"><br>
</font> </div>
<div align="center">
<table width="600" border="1" cellspacing="0" cellpadding="3" align="center" bordercolor="#000000" height="26" bordercolordark="#FFFFFF">
<tr bgcolor="#0066CC">
<td height="10" colspan="4" bgcolor="#B88230">
<div align="center"><span style="letter-spacing: 1"><font color="#FFFFFF">最新宝物消息</font></span></div>
</td>
</tr>
<tr>
<td width="165" height="9">
<div align="center"><span style="letter-spacing: 1"><font color="red">宝物名</font></span>
</div>
</td>
<td colspan="3" height="9" width="429">
<div align="center"><font color="red">增加参数</font></div>
</td>
</tr>
<!--#include file="data.asp"-->
<%
sql="SELECT * FROM 宝物 order by id desc"
Set Rs=connt.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr>
<td width="165" height="19">
<div align="center"><font color="#FF0000"><%=rs("宝物名")%></font></div>
</td>
<td colspan="3" height="19" width="429"><span style="letter-spacing: 1">
</span>
<p align="center"><span style="letter-spacing: 1"><font color="red">攻击力+<%=rs("攻击力")%>
防御力+<%=rs("防御力")%>价格+<%=rs("价格")%></font></span></p>   
</td>   
</tr>   
<%   
rs.movenext   
loop   
%>   
</table>   
<table width="600" border="1" cellspacing="0" cellpadding="3" align="center" height="26" bordercolor="#000000" bordercolordark="#FFFFFF">   
<tr>   
<td height="10" bgcolor="#B88230" colspan="4" valign="middle">   
<div align="center"><font size="2" color="#FFFFFF">最新宝物发现者</font></div>   
</td>   
</tr>   
  
<tr>   
<td width="126" height="9">   
<div align="center"><span style="letter-spacing: 1"><font color="#FF0033">宝物名</font></span>   
</div>   
</td>   
<td width="218" height="9">   
<div align="center"><font color="#FF0033">增加参数</font></div>   
</td>   
<td width="115" height="9">   
<div align="center"><span style="letter-spacing: 1"><font color="red">发现者</span></font></div>   
</td>   
<td width="131" height="9">   
<div align="center"><span style="letter-spacing: 1"><font color="red">发现时间</span></font></div>   
</td>   
</tr>   
<%   
sql="SELECT * FROM 宝物 where 拿走='1' order by id desc"   
Set Rs=connt.Execute(sql)   
do while not rs.bof and not rs.eof   
%>  
<tr>   
<td width="126" height="19">   
<div align="center"><font color="#FF0033"><%=rs("宝物名")%></font></div>  
</td>  
<td width="218" height="19">  
<div align="center"><span style="letter-spacing: 1"><font color="red">攻击力+<%=rs("攻击力")%>  
防御力+<%=rs("防御力")%> </span></div>   
</font>   
</td>   
<td width="115" height="19">   
<div align="center"><span style="letter-spacing: 1"><%=rs("获宝者")%></span></div>   
</td>   
<td width="131" height="19">   
<p align="center"><span style="letter-spacing: 1"><%=rs("时间")%></span></p>   
</td>   
</tr>   
<%   
rs.movenext   
loop   
conn.close 
set conn=nothing 
connt.close 
set connt=nothing 
%>   
</table>   
<br>   
<br>   
[ <b><a href="index.asp">登 陆 孤 岛，开 始 寻 宝</a></b> ]<br>   
<table width="551" border="1" cellspacing="1" cellpadding="0" bordercolor="#000000">   
<tr>   
<td height="94">   
<p style="margin-left: 5; margin-right: 5"><br>   
<font color="#000000">  
&lt;&lt;</font><font color="#0066CC">寻宝必读</font><font color="#000000">&gt;&gt;</font>  
</p>  
<p style="margin-left: 5; margin-right: 5; margin-bottom: 8">1、据说宝物都埋藏在另人闻风丧胆的<font color="#FF0000">江湖孤岛</font>，而且寻找到的机会可以说是微乎其微，而且随时有可能宝物没拿到手就意外死于孤岛了。<br>  
<br>  
2、宝物都是江湖中稀世珍宝，江湖黑白两道都想得到手呵，所以如果有谁拿到宝物的话可要保管好呵，因为您很快会成为黑帮的<font color="#FF0000">攻击对象</font>呵。  
</p>  
</td>  
</tr>  
</table>  
  
<br>  
</div>  
</body>  
</html>  
