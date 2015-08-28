<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=16"
Set Cn=Server.CreateObject("ADODB.Connection")
set rst=server.createobject("adodb.recordset")
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
rst.open"select 门派 from 用户 where 姓名='"& username &"' and 等级>0 and 内力>0 and 体力>0 and 银两>0",conn,1,1
if rst.bof or rst.eof then%>
<script language=vbscript>
                         MsgBox "你的条件不够啊!"
                         location.href = "javascript:window.close()"

</script>
<%
elseif trim(rst("门派"))="十八地狱" then
rst.Close                  
set rst=nothing                  
conn.Close                  
set conn=nothing
%>
<script language="vbscript">
  alert("您是鬼魂，所以您不能在这里钓鱼！")
  window.close()
</script>
<%else%>
<!--#include file="data.asp"--> 
<%
sql="select * from 钓鱼 where 姓名='"& username &"'"
Set Rs=connt.Execute(sql)
if rs.bof or rs.eof then
sql="insert into 钓鱼(姓名,时间) values ('" & username & "',now())"
connt.execute sql
Response.Redirect "diao.asp"
else
if rs("时间")<now()-1/650 then Response.Redirect "pao.asp"
end if
%>
<html>
<head>
<script language=javascript>
if(window.name!="diaole")
{ var i=1;
while (i<=50)
{
window.alert("你想作什么呀，黑我？这里是不行的，去别处玩去吧！哈！慢慢点50次！！");
i=i+1;
}
top.location.href="../../exit.asp"
}
</script>
<title>钓鱼</title>
<meta http-equiv="refresh" content="10">
<style></style>
<link rel="stylesheet" href="../css.css">
</head>

<body oncontextmenu="self.event.returnValue=false" BGCOLOR="#ffffff" text="#000000" >
      
<DIV ID="Layer1" STYLE="position:absolute; left:193px; top:180px; width:192px; height:165px; z-index:3"> 
<%
if rs("时间")>now()-1/720 then 
%>
<IMG SRC="diaoyu.gif" width="280" height="159">  
<%      
else      
%>       
<a href="diaoyuok.ASP"><IMG SRC="diaoyuok.gif" border="0" width="280" height="159"></a>     
<%     
end if      
%>       
</DIV>       
<DIV ID="Layer2" STYLE="position:absolute; left:112px; top:93px; width:252px; height:102px; z-index:1"><IMG SRC="diao1.jpg"></DIV>       
<div id="Layer3" style="position: absolute; left: 75; top: 333; width: 75; height: 12; z-index: 4"><font color=red>进行时间:<%=(Hour(time())*3600+Minute(time())*60+Second(time()))-(Hour(rs("时间"))*3600+Minute(rs("时间"))*60+Second(rs("时间")))%></font></div>   
<div id="Layer4" style="position:absolute; left:439px; top:51px; width:105px; height:66px; z-index:4"><img src="diao2.gif"></div>       
<table width="90%" border="1" cellspacing="0" cellpadding="0" height="90%" align="center" bgcolor="#FFFFFF">         
  <tr>          
    <td> </td>         
  </tr>         
</table>        
</body>          
</html>
<%
rs.Close                  
set rs=nothing                  
conn.Close                  
set conn=nothing
end if%>








