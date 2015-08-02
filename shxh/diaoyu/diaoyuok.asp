<%
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=440"
Set Cn=Server.CreateObject("ADODB.Connection")
set rst=server.createobject("adodb.recordset")
diaoyu=Application("Ba_jxqy_connstr")
Cn.Open diaoyu
randomize timer
r=int(rnd*11)
rst.open"select * from 钓鱼 where 姓名='"& username &"'",cn,1,1
if rst.bof or rst.eof then 
Response.Redirect "diaoyu.asp"
cn.close                                                                 
end if
if rst("时间")>now()-1/720 then 
Response.Redirect "diaoyu.asp"
cn.close                                                                 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("Ba_jxqy_connstr")
conn.open connstr
if r=2 or r=1 then
mess="恭喜你钓到一条大鲨鱼，可作为暗器使用，杀伤生命1500、内力1300"
sql="select * from 物品 where 名称='大鲨鱼' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof or rs.bof then
			sql="insert into 物品(名称,所有者,属性,内力,体力,价格) values ('大鲨鱼','"& username &"','暗器',-1300,-1500,2500)"
			rs=conn.execute(sql)
                        else 
				sql="update 物品 set 价格=2500,数量=数量+1  where 名称='大鲨鱼' and 所有者='" & username & "'"
				rs=conn.execute(sql)
		        end if

else
if r=3 then
mess="恭喜你钓到一条娃娃鱼，到集市卖得2万银子"
conn.execute "update 用户 set 银两=银两+20000 where 姓名='"& username &"'"
else
if r=4 or r=5 then
mess="恭喜你钓到一只大草鱼，吃下可以增加体力1500、内力1500"
sql="select * from 物品 where 名称='大草雨' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof or rs.bof then
			sql="insert into 物品(名称,所有者,属性,体力,内力,价格) values ('大草鱼','"& username &"','药品',1500,1500,3000)"
			rs=conn.execute(sql)
                        else
				sql="update 物品 set 价格=3000,数量=数量+1  where 名称='大草鱼' and 所有者='" & username & "'"
				rs=conn.execute(sql)
		        end if
else
if r=6 or r=8 then
mess="恭喜你钓到一条小鲤鱼，吃下可以增加体力500、内力500"
sql="select * from 物品 where 名称='小鲤鱼' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof or rs.bof then
			sql="insert into 物品(名称,所有者,属性,体力,内力,价格) values ('小鲤鱼','"& username &"','药品',500,500,800)"
			rs=conn.execute(sql)
                        else
				sql="update 物品 set 价格=800,数量=数量+1  where 名称='小鲤鱼' and 所有者='" & username & "'"
				rs=conn.execute(sql)
		        end if
else
if r=9 or r=10 then
mess="唉！！鱼没钓到，摔到河里体力减少900，内力减少700"
conn.execute "update 用户 set 体力=体力-900,内力=内力-700 where 姓名='" & username & "'"
else
mess="我的天啊，怎么掉上来的是只烂鞋啊，鱼钩被折断了，花了10000两银子买鱼钩，呜......"
conn.execute "update 用户 set 银两=银两-10000 where 姓名='" & username & "'"
end if
end if
end if
end if
end if
cn.execute "Delete * From 钓鱼 Where 姓名='" & username & "'"
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link href="../chat/style.css" rel=stylesheet type="text/css">
<title></title>
</head>

<body oncontextmenu=self.event.returnValue=false bgcolor="#ffffff" text="#000000">
<div align="center">
  <p>　</p>
  <table border=1 align=center cellpadding="10" cellspacing="13" height="100">
    <tr>
      <td> 
        <table>
          <tr> 
            <td valign="top"> 
              <div align="center"> 
                <p><%=mess%><br><a href="diaoyu.asp">返回继续</a></p>
              </div>
        </table>
      </td>
    </tr>
  </table>
</div>
</body>
</html>
<%                                                                 
set rs=nothing                                                                                                                                 
set conn=nothing  
%>