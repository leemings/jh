<!--#include file="../../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
my=session("yx8_mhjh_username")
if my="" then Response.Redirect "../error.asp?id=016"
if yx8_mhjh_fellow=false then%>
<script language=vbscript>
MsgBox "错误！会员功能尚未开放！"
location.href = "javascript:history.back()"
</script>
<%response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select 职业,会员 from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language="vbscript">
  alert("您不是江湖中人，我们不欢迎你！")
</script>
<%
set rs=nothing	
set conn=nothing
response.end
end if
if trim(rs("职业"))<>"郎中" then%>
<script language="vbscript">
  alert("您的职业不是江湖郎中，所以您不能在这里进行配药，请转换职业！")
location.href = "javascript:history.back()"
</script>
<%
elseif trim(rs("会员"))=false then%>
<script language="vbscript">
  alert("您不是会员，请回吧！")
location.href = "javascript:history.back()"
</script>
<%
set rs=nothing	
set conn=nothing
response.end
end if
rs.Close             
set rs=nothing             
conn.Close             
set conn=nothing  
%>
<HTML><HEAD><TITLE>配药系统</TITLE><META content="text/html; charset=gb2312" http-equiv=Content-Type><META content="MSHTML 5.00.2920.0" name=GENERATOR></HEAD><FRAMESET border=0 cols=*,170 frameBorder=no frameSpacing=0 rows=*><FRAME name=fs scrolling=no src="main.asp" frameSpacing=0 scrolling=no frameBorder=no><FRAMESET border=0 cols=* frameBorder=no frameSpacing=0 rows=50,50,0><FRAME border=0 frameBorder=no frameSpacing=0 name=ig scrolling=yes src="wupin.asp"><FRAME border=0 frameBorder=no frameSpacing=0 name=cz1 scrolling=yes src="yaopin.asp"><FRAME border=0 frameBorder=no frameSpacing=0 name=cz scrolling=no src="about:blank"></FRAMESET><FRAME SRC="UntitledFrame-2"></FRAMESET><NOFRAMES></NOFRAMES></HTML>














