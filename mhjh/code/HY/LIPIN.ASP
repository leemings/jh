<!--#include file="../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT 会员 FROM 用户 where 姓名='" &username& "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "你不是剑侠中人或者连接超时"
conn.close
response.end
else
hy=rs("会员")
if username="" then
%>
<script language=vbscript>
MsgBox "对不起，你还没有登录"
location.href = "../indexall.asp"
</script>
<%
else
if yx8_mhjh_fellow=false then
%>
<script language=vbscript>
MsgBox "错误！会员功能尚未开放！"
location.href = "javascript:history.back()"
</script>
<%
else
if hy=false then
%>
<script language=vbscript>
MsgBox "错误！你不是会员，请回吧！"
location.href = "javascript:history.back()"
</script>
<% else %>
<HTML><HEAD><TITLE>会员礼品专卖店</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK href="../style.css" rel=stylesheet>
</HEAD>
<BODY background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false topMargin="0">
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=531>
<TBODY>
<TR>
<TD  height=158 vAlign=top width="562">
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 class=mountain
width=560>
<TBODY>
<TR>
<TD vAlign=top>
<TABLE align=center border=0
cellPadding=0 cellSpacing=0 class=table1 width="100%">
<TBODY>
<TR>
<TD align=middle vAlign=center>
<P class=p9 align="center">　</P>
<P class=p9 align="center"><b><font size="4" face="隶书">会员礼品店</font></b><br>
　　这儿是江湖中最大的会员礼品店，你可以在这买到鲜花，首饰，珠宝，买了的礼物你<br>
可以在聊天室里送给你心爱的，而她用了可以增加她的美貌。<br>
<a href="top6.asp"><b><font color="#0CA339">【剑侠二十大美女排行】</font></b></a></P>
<table border=1 align=center width=550 cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">

<tr>        
<td height="15" width="79" >        
<div align="center"><font color="#000000">礼品名</font></div>        
</td>        
<td height="15" width="163" >        
<div align="center"><font color="#000000">表达情意</font></div>        
</td>        
<td height="15" width="66" >        
<div align="center"><font color="#000000">美貌</font></div>        
</td>        
<td height="15" width="66" >        
<p align="center"><font color="#000000">个性</font>        
</td>        
<td height="15" width="66" >        
<div align="center"><font color="#000000">价格</font></div>        
</td>        
<td height="15" width="55" >        
<div align="center"><font color="#000000">操 作</font></div>        
</td>        
</tr>        
<%             
sql="SELECT * FROM 礼品 where  属性='礼品'"        
Set Rs=conn.Execute(sql)        
do while not rs.bof and not rs.eof        
%>        
<tr>        
<td width="79">        
<center>               
<%=rs("名称")%>            
</center>                
<td width="163"><%=rs("表达")%></td>        
<td width="66">      
<p align="center"><%=rs("美貌")%></td>        
<td width="66">        
<p align="center"><%=rs("个性")%>　</p>        
</td>        
<td width="66">        
<center>               
<%=rs("价格")%>               
</center>        
</td>        
<td width="55">        
<center>        
<a href=lipin1.asp?id=<%=rs("id")%>>购买</a>        
</center>        
</tr>        
<%        
rs.movenext        
loop        
END IF        
END IF       
END IF       
END IF 
rs.close
set rs=nothing        
conn.close        
set conn=nothing              
%>        
</table>        
</TD></TR></TBODY></TABLE>        
<P align=center>&nbsp;</P></TD></TR></TBODY></TABLE></TD></TR>        
</TBODY></TABLE></BODY></HTML>        
