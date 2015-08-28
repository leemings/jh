<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
sj=DateDiff("s",Application("yx8_mhjh_dg"),now()) 
if sj<3 then 
s=3-sj 
Response.write "<script language='javascript'>alert ('我是魔幻之神，防止刷钱，所以无论你怎样刷\n都是无用的，请您等上["&s&"秒钟]再操作，如果不\n听劝告，你将浪费掉一个风铃！！');setTimeout('history.back();',1000);</script>"
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
rwtm=rs("任务时间")
dj=rs("任务")
yijuan=rs("风铃")
rs.close
if now()-rwtm>10 or now()-rwtm<-10 then
%>
<script language=vbscript>
MsgBox "对不起，风铃功用失效,请离开吧"
location.href = "javascript:self.close()"
</script>
<%
else
if dj<>"山谷寻宝" then
%>
<script language=vbscript>
MsgBox "奇怪了，你的任务不是" & dj & "，跑来干什么！"
location.href ="javascript:self.close()"
</script>
<%else
if username="" then%>
<script language=vbscript>
MsgBox "对不起，你还没有登录"
location.href = "../index.asp"
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
mess="辛辛苦苦打了一天木桩，没有什么收获，反而浪费了不少体力。体力因此减少了"& s &""
sql="update 用户 set 体力=体力-'"& s &"'  where 姓名='" &username& "'"
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
mess="你打木桩的姿势不对，经过大师指点，你的资质上升了"& s &"点"
sql="update 用户 set 资质=资质+'"& s &"' where 姓名='" &username& "'"
conn.execute sql
conn.close
elseif r=5 or r=6 then
mess="练功时不能专心一致，结果被大师训了一通，在思过崖面壁去了。体力因此减少200"
sql="update 用户 set 体力=体力-200 where 姓名='" &username& "'"
conn.execute sql
conn.close
elseif r=7 or r=8 then
mess="经过苦练打倒了一根木桩，结果发现木桩下面藏有葵花宝典一书，刚想拿走时被师傅收走了，看来只有等到以后再说了。体力，内力都减少了500"
sql="update 用户 set 内力=内力-500,体力=体力-500 where 姓名='" &username& "'"
conn.execute sql
conn.close
else
mess="恭喜"&username&"找到一把落怨剑，攻击8000、防御1000"
sql="select * from 物品 where 名称='落怨剑' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof and rs.bof then
			sql="insert into 物品(名称,所有者,属性,攻击,防御,数量) values ('落怨剑','" & username & "','武器',8000,1000,1)"
			conn.execute(sql)
                        else 
				sql="update 物品 set 数量=数量+1 where 名称='落怨剑' and 所有者='" & username & "'"
				conn.execute(sql)
		        end if
                        rs.close
                        set rs=nothing
end if
set conn=nothing
%>
<script language=vbscript>
MsgBox "<%=mess%>"
location.href = "d2.asp"
</script>

<%end if
end if
end if
end if%>
