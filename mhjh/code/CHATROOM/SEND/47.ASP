<%
function xindong(un,gr,mg)
if gr<10 then
Response.Write "<script language=JavaScript>{alert('心动感觉，需要等级10级，你不够呀！！');}</script>"
Response.End
end if
rst.Open "select 内力,会员 from 用户 where 姓名='"&un&"'",conn
if rst("内力")<100000  then
Response.Write "<script language=JavaScript>{alert('内力不够100000，不能使用本操作！！');}</script>"
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
conn.execute "update 用户 set 内力=内力-100000 where 姓名='"&un&"'"
xindong="<marquee width=100% behavior=alternate scrollamount=10><font color=red><font face=Webdings class=p>&yuml;</font>〖心动感觉〗</font>" & mg & "(" & un & ")" & "</marquee>"
rst.close
set rst=nothing
end function
%>