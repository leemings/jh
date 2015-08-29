<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT id FROM 用户 WHERE  姓名='" & sjjh_name &"'",conn,2,2
sjjh_id=rs("id")
rs.close
rs.open "Select  * from 投票",conn,3,3
ks=rs("开始时间")
jz=rs("截止时间")
dj=rs("等级")
rs.close
rs.open "Select * from 投票者 where 投票ID="&sjjh_id,conn,3,3
if rs.eof or rs.bof then
	tp=0
else
	tp=1
end if
rs.close
sql="Select count(*) As 数量 from 投票者"
set tmprs=conn.execute(sql)
ytrs=tmprs("数量")
tmprs.close
set tmprs=nothing
sql="Select count(*) As 数量 from 候选人"
set tmprs=conn.execute(sql)
hxr=tmprs("数量")
tmprs.close
set tmprs=nothing
sql="select * from 候选人 order by 得票数 desc"
set rs=conn.execute(sql)
%>
<html>
<head>
<title>候选人管理♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../css.css" rel=stylesheet>
<script language="JavaScript">
function load()
{var name=navigator.appName
var vers=navigator.appVersion
if(name=="Netscape")
{window.location.reload()
}else
{history.go(0)
}}</script>
</head>
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="10" valign="bottom"><img src="../images/tp1.gif" width="96" height="23"></td>
<td width="768"><img src="../images/tp4.jpg" width="158" height="100"></td>
<td width="768">&nbsp;</td>
</tr>
<tr bgcolor="#000000">
<td width="10">&nbsp;</td>
<td width="768"><img src="../images/tp3.jpg" width="158" height="60"></td>
<td width="768" valign="bottom" align="right"><img src="../images/tp2.gif" width="80" height="18"></td>
</tr>
</table>
<table border="0" cellpadding="4" cellspacing="0" width="775" height="5">
<tr>
<td width="100%" height="1"> 【投票结果】 </td>
</tr>
<tr>
<td width="100%" height="19"> 现在是：<%=now()%> </td>
</tr>
</table>
<blockquote>
<p align="left"><b>[功能说明]</b>　<br>
投票时间：<font color=red><%=ks%></font> 至 <font color=red><%=jz%><br>           
</font>目前共有候选人：<font color=red><%=hxr%></font> 位；已投票：<font color=red><%=ytrs%></font>           
票。<a href="javascript:load()">【刷新】</a><a href="javascript:history.go(-1)">【返回】</a>           
</p>
</blockquote>
<form method="post" action="pollviewpoll.asp">
<table border="1" align="center" cellpadding="4" cellspacing="0" width="726" bordercolor="#800000">
<tr>
      <td align="center" width="31" bordercolorlight="#800000"><font color="#FF6600">序 
        号</font></td>
      <td align="center" width="56" bordercolorlight="#800000"><font color="#FF6600">候 
        选 人 </font></td>
      <td align="center" width="427" bordercolorlight="#800000"><font color="#FF6600">得 
        票 数</font></td>
      <td align="center" width="51" bordercolorlight="#800000"><font color="#FF6600">得 
        票 率 </font></td>
      <td align="center" width="36" bordercolorlight="#800000"><font color="#FF6600">名 
        次</font></td>
      <td align="center" width="63" bordercolorlight="#800000"><font color="#FF6600">投 
        票</font></td>
</tr>
<%i=1
j=1
if not(rs.eof or rs.bof) then sy=rs("得票数")
	do while not (rs.eof or rs.bof)%>
<tr>
      <td align="center" width="31" bordercolorlight="#800000" height="2"><%=i%></td>
      <td width="56" bordercolorlight="#800000" height="2"> 
        <div align="center"><%=rs("姓名")%></div></td>
      <td nowrap width="427" bordercolorlight="#800000" height="2"> 
        <%if rs("得票数")>0 then
        n=int(rs("得票数")/ytrs*100)
        %>
        <img src="poll.gif" width="<%=n*4%>" height="18"> 
        <%end if%>
        &nbsp;<%=rs("得票数")%></td>           
      <td align="right" width="51" bordercolorlight="#800000" height="2"> 
        <%if ytrs=0 then%>
        <div align="center">0%            
          <%else
	n=int(rs("得票数")/ytrs*100)
	%>
          <%=n%>%            
          <%end if%>
        </div>
      </td>
      <td align="center" width="36" bordercolorlight="#800000" height="2"> 
        <%if rs("得票数")=sy then
	j=j-1
end if
if j=0 then
	j=1
end if
%>
        <%=j%> </td>
        
       <%yxtp=1
if sjjh_jhdj<dj then
	yxtp=0
end if
if tp<>0 then
	yxtp=0
end if
if now()<ks then
	yxtp=0
end if
if now()>jz then
	yxtp=0
end if%>
      <td align="center" width="63" bordercolorlight="#800000" height="2"> 
        <%if yxtp=1 then%>
        <a href="pollviewpoll.asp?id=<%=rs("姓名")%>"><font color="#0000FF"><u>投他一票</u></font></a> 
        <%else%>
        不能投票 
        <%end if%>
      </td>
</tr>
<%
	sy=rs("得票数")
	j=j+1
	i=i+1
	rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<tr>
      <td align="center" colspan="6"> 
        <%yxtp=1
if sjjh_jhdj<dj then
	yxtp=0%>
        您的战斗等级没有达到<%=dj%>级，不能投票! 
        <%end if
if tp<>0 then
	yxtp=0%>
        您已经投过票了，不能重复投票 
        <%end if
if now()<ks then
	yxtp=0%>
        投票活动尚未开始，不能投票 
        <%end if
if now()>jz then
	yxtp=0%>
        投票活动已经结束，不能投票 
        <%end if%>
      </td>
</tr>
</table>
</form>
</body>
</html>

