<%
function nuhou(un,st,gr,mg)
if gr<15 then
Response.Write "<script language=JavaScript>{alert('等级不够15，不能操作！！');}</script>"
Response.End
end if
rst.Open "select 内力,会员 from 用户 where 姓名='"&un&"'",conn
if rst("内力")<500000  then
Response.Write "<script language=JavaScript>{alert('内力不够500000，不能使用本操作！！');}</script>"
Response.End
end if
if rst("会员")<>true then
Response.Write "<script language=JavaScript>{alert('你不是会员，不能使用本操作！！');}</script>"
rst.close
set rst=nothing
conn.close
set conn=nothing
Response.End
end if
conn.execute "update 用户 set 内力=内力-500000 where 姓名='"&un&"'" 
nuhou="<marquee height=50 behavior=alternate loop=200 direction=up onmouseover=this.stop(); onmouseout=this.start();><marquee behav onmouseover=this.stop(); onmouseout=this.start();><ior=alternate onmouseover=this.stop(); onmouseout=this.start();>〖狂狮怒吼〗"&un&"对"&st&"说:" & mg & "" & "</marquee></marquee>"
rst.close
set rst=nothing
end function
%>
