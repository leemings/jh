<%
function hmen(un,st,co)
rst.Open "select 结盟 from 门派 where 门派='"&co&"'",conn
guojia=rst("结盟")
rst.Close
rst.Open "select 门派 from 用户 where 姓名='"&st&"'",conn
men=rst("门派")
rst.Close
rst.Open "select 结盟 from 门派 where 门派='"&men&"'",conn
jimen=rst("结盟")
rst.Close
rst.Open "select 身份,间接3 from 用户 where 姓名='"&un&"'",conn
if rst("身份")="掌门" then
if jimen="无" and  rst("间接3")=""&men&"" then
conn.execute "update 用户 set 间接3='无' where 姓名='"&un&"'" 
hmen="<font color=FF0000>【毁盟】</font>"&co&"掌门##觉得"&men&"掌门%%太不给面子!不想再理"&men&"的人啦!"
else
if guojia="" or guojia="无" then
response.write "<script language='javascript'>alert('你的门派没和别人结盟啊!');</script>"
response.end
else
if guojia=""&men&"" and jimen=""&co&"" then
conn.execute "update 用户 set 间接3='无' where 姓名='"&un&"'"
conn.execute "update 用户 set 间接3='无' where 姓名='"&st&"'"
conn.execute "update 门派 set 结盟='无' where 门派='"&co&"'"
conn.execute "update 门派 set 结盟='无' where 门派='"&men&"'"
hmen="<font color=FF0000>【毁盟】</font>"&co&"掌门##觉得"&men&"不够资格和本门结盟断然和"&men&"断交!"
else
hmen="<font color=FF0000>【毁盟】</font>"&co&"掌门##你搞错啦,"&men&"不是你盟友嘛!"
end if
end if
end if
else
hmen="<font color=FF0000>【毁盟】</font>"&co&"的##你不是掌门,是想造反吗!" 
end if
rst.Close
end function
%>