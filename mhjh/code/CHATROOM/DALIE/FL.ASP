<!--#include file="data.asp"-->
<!--#include file="dalie.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
nickname=session("yx8_mhjh_username")
sql="select * from 打猎 where dalie<>'无'"
set rs=connt.execute(sql)
if rs.eof then
%>
<script>
alert('现在还没有猎物可以打，请少后再来把。');
window.close();
</script>
<%
else
%>
<%
sql1="update 用户 set 内力=内力-250,体力=体力-250,银两=银两-1000 where 姓名='"&nickname&"'"
conn.execute sql1
%>
<html>
<title>打猎</title>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<script>
function dalie()
{
location.href='da.asp';
}
</script>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="444" height="278">
  <param name="Movie" value="dalie.swf">
  <param name="Src" value="dalie.swf">
  <param name="WMode" value="Window">
  <param name="Play" value="-1">
  <param name="Loop" value="-1">
  <param name="Quality" value="High">
  <param name="SAlign" value>
  <param name="Menu" value="-1">
  <param name="Base" value>
  <param name="Scale" value="ShowAll">
  <param name="DeviceFont" value="0">
  <param name="EmbedMovie" value="0">
  <param name="BGColor" value><embed src="dalie.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="444" height="278"> 
</object>
</body> 
</html>
<%end if%>
