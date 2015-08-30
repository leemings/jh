<%
if Session("sjjh_Name")="" then
%>
<script language=vbscript>
MsgBox "你不是还没有登陆或者超时连接呀！"
location.href = "javascript:self.close()"
</script>
<%end if%>

<%
set conn=server.createobject("adodb.connection") 
conn.open Application("sjjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
name=Session("sjjh_Name")
sql="select * from 用户 where 姓名='" & name & "'"
set rs=conn.execute(sql)
if rs("银两")<88888 then
%>
<script language=vbscript>
MsgBox "我倒，连88888两银两都没有呀！"
location.href = "javascript:self.close()"
</script>
<% end if%>
<%
sql="update 用户 set 银两=银两-88888 where 姓名='" & name & "'"
conn.execute sql
%>
<html>
<head>
<title>快乐江湖-宝物调查所</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../setup.css">
<style type="text/css">
BODY {
scrollbar-face-color:#efefef; 
scrollbar-shadow-color:#000000; 
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#efefef;
scrollbar-darkshadow-color:#efefef;
scrollbar-track-color:#efefef;
scrollbar-arrow-color:#000000;
}
</style>
</head>

<body  background='bg.gif' oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5">
<div align="center"><font color="#FFFFFF"> </font> <font size="2"> <font color="#800000">[你花了88888两来到了快乐江湖宝物调查所，这里有各种都市人士从不同的地方打探到的藏宝消息]</font><br>
</font><font size="2"><br>
  </font> <font color="#008080" size="4"><b>孤 岛 冒 险</b></font><font size="2"> 
  </font> </div>
<div align="center">　 </div>
<div align="center">
<table width="600" border="1" cellspacing="0" cellpadding="3" align="center" bordercolor="#000000" height="26" bordercolordark="#FFFFFF">
<tr bgcolor="#0066CC">
<td height="10" colspan="4" bgcolor="#B88230">
<div align="center"><span style="letter-spacing: 1"><font color="#FFFFFF">最新宝物消息</font></span></div>
</td>
</tr>
<tr>
<td width="165" height="9" bgcolor="#008080">
<div align="center"><span style="letter-spacing: 1"><font color="#FFFF00">宝物名</font></span>
</div>
</td>
<td colspan="3" height="9" width="429" bgcolor="#008080">
<div align="center"><font color="#FFFF00">增加参数</font></div>
</td>
</tr>
<%
sql="SELECT * FROM 宝物 order by id desc"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr>
<td width="165" height="19" bgcolor="#008080">
<div align="center"><font color="#FFFF00"><%=rs("宝物名")%></font></div>
</td>
<td colspan="3" height="19" width="429" bgcolor="#008080">
<p align="center"><span style="letter-spacing: 1"><font color="#FFFF00">攻击力+<%=rs("攻击力")%>
防御力+<%=rs("防御力")%> </font></span></p>    
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
<td width="126" height="9" bgcolor="#008080">  
<div align="center"><span style="letter-spacing: 1"><font color="#FFFF00">宝物名</font></span>  
</div>  
</td>  
<td width="218" height="9" bgcolor="#008080">  
<div align="center"><font color="#FFFF00">增加参数</font></div>  
</td>  
<td width="115" height="9" bgcolor="#008080">  
<div align="center"><span style="letter-spacing: 1"><font color="#FFFF00">发现者</font></span></div>  
</td>  
<td width="131" height="9" bgcolor="#008080">  
<div align="center"><span style="letter-spacing: 1"><font color="#FFFF00">发现时间</font></span></div>  
</td>  
</tr>  
<%  
sql="SELECT * FROM 宝物 where 拿走='1' order by id desc"  
Set Rs=conn.Execute(sql)  
do while not rs.bof and not rs.eof  
%> 
<tr>  
<td width="126" height="19" bgcolor="#008080">  
<div align="center"><font color="#FFFF00"><%=rs("宝物名")%></font></div> 
</td> 
<td width="218" height="19" bgcolor="#008080"> 
<div align="center"><span style="letter-spacing: 1"><font color="#FFFF00">攻击力+<%=rs("攻击力")%> 
防御力+<%=rs("防御力")%> </font> </span></div>    
</td>  
<td width="115" height="19" bgcolor="#008080">  
<div align="center"><span style="letter-spacing: 1"><%=rs("获宝者")%></span></div>  
</td>  
<td width="131" height="19" bgcolor="#008080">  
<p align="center"><span style="letter-spacing: 1"><%=rs("时间")%></span></p>  
</td>  
</tr>  
<%  
rs.movenext  
loop  
conn.close  
%>  
</table>  
<br>  
<br> 
[ <b><a href="../index.asp"><font color="#FF0000">登 陆 孤 岛，开 始 寻 宝</font></a></b> <font color="#FF0000">  
</font> ]<br>  
  <table width="600" border="1" cellspacing="1" cellpadding="0" bordercolor="#000000">
    <tr>  
<td height="94" bgcolor="#008080">  
<p style="margin-left: 5; margin-right: 5"><br>  
<font color="#FFFF00"> 
&lt;&lt;寻宝必读&gt;&gt;</font> 
</p> 
        <p style="margin-left: 5; margin-right: 5; margin-bottom: 8"><font color="#FFFFFF">1、据说宝物都埋藏在另人闻风丧胆的</font><font color="#FFFF00">[E线孤岛]</font><font color="#FFFFFF">，而且寻找到的机会可以说是微乎其微，而且随时有可能宝物没拿到手就意外死于孤岛了。</font><br> 
<br> 
<font color="#FFFFFF"> 
2、宝物都是江湖中稀世珍宝，江湖黑白两道都想得到手呵，所以如果有谁拿到宝物的话可要保管好呵，因为您很快会成为帮派的</font><font color="#FFFF00">攻击对象</font><font color="#FFFFFF">呵。</font> 
</p> 
</td> 
</tr> 
</table> 
 
<br> 
</div> 
</body> 
</html>
