<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if session("yx8_mhjh_inchat")="" then
Response.write "<script language='javascript'>alert ('你不能进入,请先进入聊天室再来。谢谢合作');setTimeout('history.back();',1000);</script>"
Response.End 
end if
sj=DateDiff("s",Application("yx8_mhjh_dg"),now()) 
if sj<5 then 
s=5-sj 
Response.write "<script language='javascript'>alert ('我是剑侠之神，请您等上["&s&"秒钟]再操作！');setTimeout('history.back();',1000);</script>"
Response.End 
end if 
Application.Lock 
Application("yx8_mhjh_dg")=now() 
Application.UnLock 
username=session("yx8_mhjh_username")
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT 银两,存款 FROM 用户 where 姓名='" &username& "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "你连接超时"
conn.close
response.end
else
mymoney=rs("银两")
cunk=rs("存款")
if username="" then
%>
<script language=vbscript>
MsgBox "对不起，你还没有登录"
location.href = "../exit.asp"
</script>
<%
else
if mymoney>30000 then
%>
<script language=vbscript>
MsgBox "错误！你目前银两大于30000,官府不放粮了！"
location.href = "javascript:history.back()"
</script>
<%
else
if cunk>30000 then
%>
<script language=vbscript>
MsgBox "错误！你目前存款大于30000,官府不放粮了！"
location.href = "javascript:history.back()"
</script>
<%
else
randomize timer
r=int(rnd*50)
s=1
nl=0
sm=0
nu=int(rnd*18)+1
s=int(rnd*100)
sql="update 用户 set 银两=银两+'"& s &"'  where 姓名='" &username& "'"
conn.execute sql
%>
<script language=vbscript>
MsgBox "<%=my%>"& "恭喜您!获得"&"<%=s%>"&"两白银！"
location.href="javascript:top.window.close();"
</script>
<%end if
end if
end if
end if
rs.Close             
set rs=nothing             
conn.Close             
set conn=nothing       
%>


