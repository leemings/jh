<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
sj=DateDiff("s",Application("yx8_mhjh_dg"),now()) 
if sj<3 then 
s=3-sj 
Response.write "<script language='javascript'>alert ('我是魔幻侠之神，防止刷钱，所以无论你怎样刷\n都是无用的，请您等上["&s&"秒钟]再操作，如果不\n听劝告，你将浪费掉一个风铃！');setTimeout('history.back();',1000);</script>"
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
mess="恭喜"&username&"找到一把狂刀，杀伤体力1200、内力900"
sql="select * from 物品 where 名称='黑血神针' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof and rs.bof then
			sql="insert into 物品(名称,所有者,属性,攻击,防御,数量) values ('黑血神针','" & username & "','暗器',1200,900,1)"
			conn.execute(sql)
                        else 
				sql="update 物品 set 数量=数量+1 where 名称='黑血神针' and 所有者='" & username & "'"
				conn.execute(sql)
		        end if
                        rs.close
                        set rs=nothing
elseif r=3 then
randomize timer
r=int(rnd*50)
s=1
nl=0
sm=0
nu=int(rnd*18)+1
s=int(rnd*100)
mess="你刚伸手准备偷窃的时候，被大师发现，结果被痛打，损失银两1000"
sql="update 用户 set 银两=银两-1000 where 姓名='" &username& "'"
conn.execute sql
conn.close
elseif r=5 or r=6 then
mess="费了半天工夫什么也没有得到，白白浪费了2500的体力。"
sql="update 用户 set 体力=体力-2500 where 姓名='" &username& "'"
conn.execute sql
conn.close
elseif r=7 or r=8 then
mess="原来大师袖口藏着少林绝学易经经，听说看后可以提高所有数值，刚想占有又被发现，失去内力1500，体力2000……惨！"
sql="update 用户 set 内力=内力-1500,体力=体力-2000 where 姓名='" &username& "'"
conn.execute sql
conn.close
else
mess="恭喜"&username&"找到一把袖口剑，攻击1000、防御1000"
sql="select * from 物品 where 名称='袖口剑' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof and rs.bof then
			sql="insert into 物品(名称,所有者,属性,攻击,防御,数量) values ('袖口剑','" & username & "','武器',1000,1000,1)"
			conn.execute(sql)
                        else 
				sql="update 物品 set 数量=数量+1 where 名称='袖口剑' and 所有者='" & username & "'"
				conn.execute(sql)
		        end if
                        rs.close
                        set rs=nothing
end if
set conn=nothing
%>
<script language=vbscript>
MsgBox "<%=mess%>"
location.href = "a.asp"
</script>
<%end if
end if
end if
end if%>
