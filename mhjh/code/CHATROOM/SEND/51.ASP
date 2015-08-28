<%
function xintiao(un,gr,mg)
if gr<20 then
Response.Write "<script language=JavaScript>{alert('此操作需要等级20，才可以操作！！');}</script>"
Response.End
end if
rst.Open "select 内力,会员 from 用户 where 姓名='"&un&"'",conn
if rst("内力")<800000  then
Response.Write "<script language=JavaScript>{alert('内力不够800000，不能使用本操作！！');}</script>"
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
conn.execute "update 用户 set 内力=内力-800000 where 姓名='"&un&"'" 
xintiao="<marquee height=50 behavior=alternate loop=200 direction=up><font color=red><font face=Webdings class=p>&yuml;</font>〖心跳心语〗</font>" & mg & "(" & un & ")" & "</marquee>"
rst.close
set rst=nothing	
end function
%>