<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
sj=DateDiff("s",Application("yx8_mhjh_dg"),now()) 
if sj<5 then 
s=5-sj 
Response.write "<script language='javascript'>alert ('我是剑侠之神，防止刷钱，所以无论你怎样刷\n都是无用的，请您等上["&s&"秒钟]再操作！');setTimeout('history.back();',1000);</script>"
Response.End 
end if 
Application.Lock 
Application("yx8_mhjh_dg")=now() 
Application.UnLock 
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 用户 where 姓名='" &username& "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "你不是剑侠中人或者连接超时"
conn.close
response.end
else
mymoney=rs("任务")
dj=rs("等级")
rs.close
if username="" then
%>
<script language=vbscript>
MsgBox "对不起，你还没有登录"
location.href = "../index.asp"
</script>
<%
else
if mymoney="" then
%>
<script language=vbscript>
MsgBox "奇怪了，你还没有到稻香村村长那里接受任务，跑来干什么！"
location.href = "javascript:history.back()"
</script>
<%
else
if dj<10 then
%>
<script language=vbscript>
MsgBox "错误！你目前的等级低于10,不能做任务，请耐心等待升级吧……"
location.href = "javascript:history.back()"
</script>
<%
else
randomize timer
r=int(rnd*10)

if r<=2 then
randomize timer
r=int(rnd*50)
s=1
nl=0
sm=0
nu=int(rnd*18)+1
s=int(rnd*100)
mess="南太行山的野狼果然厉害，你被群狼攻击，身负重伤。你的攻击和防御因此减少了"& s &""
sql="update 用户 set 攻击=攻击-'"& s &"',防御=防御-'"& s &"'  where 姓名='" &username& "'"
conn.execute sql
conn.close
elseif r=3 then
randomize timer
r=int(rnd*50)
s=1
nl=0
sm=0
nu=int(rnd*18)+1
s=int(rnd*100)
mess="你在南太行山遇到了山寨精英，被掠夺去了银子"& s &"两"
sql="update 用户 set 银两=银两-'"& s &"'  where 姓名='" &username& "'"
conn.execute sql
conn.close
elseif r=5 or r=6 then
mess="南太行山险道你见一剑客练剑，一打听原来是令狐冲，悄悄偷学了几招，资质增加了20"
sql="update 用户 set 资质=资质+20  where 姓名='" &username& "'"
conn.execute sql
conn.close
elseif r=7 or r=8 then
mess="经过与群狼战斗，虽然体力损失10000，但你得到了狼皮一张，快到西安城找孙周换华山玉佩吧！"
sql="update 用户 set 任务='狼皮',体力=体力-10000  where 姓名='" &username& "'"
conn.execute sql
conn.close
else
mess="恭喜您!你来到南太行山山寨竟然没有遇到任何人，直接从山寨头子宝座上得到了风铃佩，可以到苏州城找月芽儿换勾践枪了！"
sql="update 用户 set 任务='风铃佩'  where 姓名='" &username& "'"
conn.execute sql
conn.close
end if
set conn=nothing
%>
<script language=vbscript>
MsgBox "<%=mess%>"
location.href = "taihang.asp"
</script>

<%end if
end if
end if
end if%>
