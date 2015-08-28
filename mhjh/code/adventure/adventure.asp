<!--#include file="../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
my=session("yx8_mhjh_username")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select 体力,任务 from 用户 where 姓名='" & my & "'"
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
set rs=nothing	
set conn=nothing
if session("yx8_mhjh_userposr")="" then
session("yx8_mhjh_userposr")=1
session("yx8_mhjh_userposc")=0
session("yx8_mhjh_userfight")="none"
end if
%>
<head>
<title></title>
</head>
<frameset cols="539,*,110" border=0 frameborder="0" framespacing="0">
<frameset rows="*,344" border=0>
<frame name=mapfrm src="map.asp"  scrolling=no>
<frame name=msgfrm src='standby.htm' scrolling=auto noresize>
</frameset>
<frameset rows="200,380,0,110"  border=0>
<frame name=behfrm src="about:blank" scrolling=no>
<frame name=confrm src="about:blank" scrolling=no>
<frame name=actfrm src="about:blank">
<frame name=optfrm src="welcome.asp" scrolling=auto marginwidth="0" marginheight="0">
</frameset>
<frame name=mhjjy src='1.htm' scrolling=no>
</frameset>
</frameset>