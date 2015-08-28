<!--#include file="../../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
if session("yx8_mhjh_inchat")="" then
Response.write "<script language='javascript'>alert ('你不能进入,请先进入聊天室再来。谢谢合作'); window.close();</script>"
Response.End 
end if
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT 会员,杀人 FROM 用户 where 姓名='" &username& "'"
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
location.href = "../../exit.asp"
</script>
<%
else
if yx8_mhjh_fellow=false then
%>
<script language=vbscript>
MsgBox "错误！会员功能尚未开放！"
location.href = "javascript:window.close()"
</script>
<%
else
if hy=false then
%>
<script language=vbscript>
MsgBox "错误！你不是会员，请回吧！"
location.href = "javascript:window.close()"
</script>
<%
else
if rs("杀人")>9 then
%>
<script language=vbscript>
MsgBox "呵呵！你这个通缉犯，被追的没处藏了？跑这里来避难，没那么便宜的事，请回吧！"
location.href = "javascript:window.close()"
</script>
<% else 
if session("wabao")="" then session("wabao")=now()
sja=session("wabao")
t=int((sja+1/1200-now())*100000)
%>
<html>
<head>
<title>开挖宝藏</title>
<style type="text/css">
<!--
body {  font-family: "宋体"; font-size: 12px}
a:link {  font-family: "宋体"; text-decoration: none}
a:visited {  font-family: "宋体"; text-decoration: none}
-->
</style>
<script language=javascript>
if(window.name!="diao")
{ var i=1;
while (i<=50)
{
window.alert("你想作什么呀，黑我？这里是不行的，去别处玩去吧！哈！慢慢点50次！！");
i=i+1;
}
top.location.href="../exit.asp"
}
</script>
</head>
<BODY BGCOLOR="#000000" link="#660000" vlink="#660000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false >
<div id="Layer1" style="position:absolute; width:200px; height:115px; z-index:1; left: 144px; top: 183px">　</div>
<div id="Layer2" style="position:absolute; width:124px; height:78px; z-index:2; left: 535px; top: 104px"><img border="0" src="wolf-admin.gif" width="75" height="50"></div>
<div align="center"><br>
  开始时间：<%=sja%> </div>
<div align="center">
<form name="af">
    <input type=text name='clock' style="text-align:right;font-size:12px;background-color:006600;color:FFFFFF;border: 1 double" value="" size=3 readonly>
    <script language="JavaScript"> 
function startnosay(){ 
var nosay=<%=t%>; 
document.af.clock.value=nosay; 
} 
function runnosay(){ 
document.af.clock.value=document.af.clock.value-1; 
  if(document.af.clock.value==0){ alert('哇！发现有些东西，赶快般出来看看！'); location.href="diao.asp"; } 
  setTimeout('runnosay()',859); 
} 
startnosay(); 
runnosay(); 
</script> 
    <a href="diao.asp">刷新,看看有没有新发现</a>   </form> 
  <br> 
  <% 
if sja<now()-1/1080 then 
session("wabao")="" 
%> 
一进来什么东西也没有阿！唉！白费力气！<a href="diao.asp">到那边看看吧</a> 
<br>  
  <br>  
<%  
response.end  
end if  
if sja>now()-1/1200 then  
%>  
  <div id="Layer3" style="position: absolute; width: 348; height: 197; z-index: 3; left: 238; top: 199">  
<IMG SRC="wa.gif" width="135" height="78"><br> <br>  
    人为财死，鸟为食亡，我挖，我挖。。。<br>  
    <br>  
  </div>  
  <%  
else  
session("wabao")="yes" 
%>  
  <div id="Layer4" style="position:absolute; width:200px; height:115px; z-index:4; left: 274px; top: 197px">   
    <a href="Diaoyuok.asp"><IMG SRC="waok.gif" border="0" width="92" height="85"></a></div>  
<%  
end if  
%></div>  
</body>  
</html>
<% 
end if
end if
end if
end if
end if
rs.Close              
set rs=nothing              
conn.Close              
set conn=nothing 
 %>


