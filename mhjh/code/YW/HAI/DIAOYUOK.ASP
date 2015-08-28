<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.open"select id from 用户 where 姓名='"& username &"' and 等级>10 and 内力>0 and 体力>0 and 银两>0",conn,1,1
if rst.bof or rst.eof then
set conn=nothing  
%>
<script language=vbscript>
                         MsgBox "你的条件不够啊!"
                         location.href = "javascript:window.close()"
</script>
<%
else
rst.close
randomize timer
r=int(rnd*16)
%>
<!--#include file="data.asp"--><%
sql="select * from 钓鱼 where 姓名='"& username &"'"
 Set Rs=connt.Execute(sql)
if rs.bof or rs.eof then 
Response.Redirect "diaoyu.asp"
connt.close                                                                 
end if
if rs("时间")>now()-1/720 then 
Response.Redirect "diao.asp"
connt.close                                                                 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
if r=2 or r=1 then
mess="恭喜你钓到一条大鲨鱼，可作为暗器使用，攻击500000"
sql="select * from 物品 where 名称='大鲨鱼' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof or rs.bof then
			sql="insert into 物品(名称,属性,体力,内力,攻击,防御,特效,价格,数量,所有者,寄售,装备) values ('大鲨鱼','暗器',0,0,500000,0,'无',600000,1,'"& username &"',False,False)"
			rs=conn.execute(sql)
                        else 
				sql="update 物品 set 价格=600000,数量=数量+1  where 名称='大鲨鱼' and 所有者='" & username & "'"
				rs=conn.execute(sql)
		        end if

else
if r=3 then
mess="恭喜你钓到一条娃娃鱼，到集市卖得100000银子"
conn.execute "update 用户 set 银两=银两+100000 where 姓名='"& username &"'"
else
if r=4 or r=5 then
mess="恭喜你钓到一只大草鱼，吃下可以增加体力250000、内力100000"
sql="select * from 物品 where 名称='大草鱼' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof or rs.bof then
			sql="insert into 物品(名称,属性,体力,内力,攻击,防御,价格,特效,数量,所有者,寄售,装备) values ('大草鱼','药品',250000,100000,0,0,1000000,'无',1,'"& username &"',False,False)"
			rs=conn.execute(sql)
                        else
				sql="update 物品 set 价格=1000000,数量=数量+1  where 名称='大草鱼' and 所有者='" & username & "'"
				rs=conn.execute(sql)
		        end if
else
if r=6 or r=8 then
mess="恭喜你钓到一条小鲤鱼，吃下可以增加体力300000、内力200000"
sql="select * from 物品 where 名称='小鲤鱼' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof or rs.bof then
			sql="insert into 物品(名称,属性,体力,内力,攻击,防御,价格,特效,数量,所有者,寄售,装备) values ('小鲤鱼','药品',300000,200000,0,0,800000,'无',1,'"& username &"',False,False)"
			rs=conn.execute(sql)
                        else
				sql="update 物品 set 价格=800000,数量=数量+1  where 名称='小鲤鱼' and 所有者='" & username & "'"
				rs=conn.execute(sql)
		        end if
else
if r=9 or r=10 then
mess="唉！！鱼没钓到，摔到河里体力减少50000，内力减少20000"
conn.execute "update 用户 set 体力=体力-50000,内力=内力-20000 where 姓名='" & username & "'"
else
mess="我的天啊，怎么掉上来的是只烂鞋啊，鱼钩被折断了，花了100000两银子买鱼钩，呜......"
conn.execute "update 用户 set 银两=银两-100000 where 姓名='" & username & "'"
end if
end if
end if
end if
end if
%>
<!--#include file="data.asp"--><%
sql="Delete * From 钓鱼 Where 姓名='" & username & "'"
connt.execute sql
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link href="../css.css" rel=stylesheet type="text/css">
<title></title>
</head>

<body oncontextmenu=self.event.returnValue=false  background='<%=bgimage%>'>
<div align="center">
  <p>　</p>
  <table border=1 align=center cellpadding="10" cellspacing="13" height="100">
    <tr>
      <td> 
        <table>
          <tr> 
            <td valign="top"> 
              <div align="center"> 
                <p><%=mess%><br><a href="diao.asp">返回继续</a></p>
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
set rst=nothing
set connt=nothing  
end if
%>