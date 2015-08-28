<!--#include file="../../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
my=session("yx8_mhjh_username")
if my="" then Response.Redirect "../error.asp?id=016"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select 门派,道德 from 用户 where 姓名='" & my & "'"
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
if trim(rs("门派"))="十八地狱" then%>
<script language="vbscript">
  alert("您现在还是鬼魂呢，请尽快去凤凰蘖盘那进化成人吧！")
location.href = "javascript:history.back()"
</script>
<%
elseif trim(rs("道德"))<0 then%>
<script language="vbscript">
  alert("你的道德太少，这里不欢迎缺乏道德的人！")
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
<HTML><HEAD><TITLE>铸造武器装备</TITLE><META content="text/html; charset=gb2312" http-equiv=Content-Type><META content="MSHTML 5.00.2920.0" name=GENERATOR></HEAD><FRAMESET border=0 cols=*,170 frameBorder=no frameSpacing=0 rows=*><FRAME name=fs scrolling=no src="main.asp" frameSpacing=0 scrolling=no frameBorder=no><FRAMESET border=0 cols=* frameBorder=no frameSpacing=0 rows=600,50><FRAME border=0 frameBorder=no frameSpacing=0 name=ig scrolling=yes src="wupin.asp"><FRAME border=0 frameBorder=no frameSpacing=0 name=cz scrolling=no src="about:blank"></FRAMESET><NOFRAMES></NOFRAMES></HTML>














