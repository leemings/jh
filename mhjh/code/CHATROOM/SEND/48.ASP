<%
function jmen(un,st,co)
if un=st then 
response.write "<script language='javascript'>alert('不能和自己结盟');</script>"
response.end
else
rst.Open "select 国家 from 门派 where 门派='"&co&"'",conn
guojia=rst("国家")
rst.Close
rst.Open "select 门派 from 用户 where 姓名='"&st&"'",conn
men=rst("门派")
rst.Close
rst.Open "select 国家,结盟 from 门派 where 门派='"&men&"'",conn
guo=rst("国家")
jimen=rst("结盟")
rst.Close
if guojia<>guo then
response.write "<script language='javascript'>alert('你竟敢和蛮子勾结,想遗臭万年吗?掌门所在国家不同，无法结盟！');</script>"
response.end
else
if co=jimen then
response.write "<script language='javascript'>alert('你们二派已经结盟了啊!');</script>"
response.end
else
rst.Open "select 身份,间接3 from 用户 where 姓名='"&un&"'",conn
if rst("身份")="掌门" then
if rst("间接3")="" or rst("间接3")="无" then
conn.execute "update 用户 set 间接3='"&men&"' where 姓名='"&un&"'"
jmen="<font color=FF0000>【歃血为盟】</font>"&co&"掌门##想和"&men&"掌门%%结为同盟!共抗外敌!"
else
ty=rst("间接3")
if rst("间接3")=""&men&"" then
rst.Close
rst.Open "select * from 用户 where 姓名='"&st&"'",conn
if rst("间接3")=""&co&"" then
conn.execute "update 门派 set 结盟='"&men&"' where 门派='"&co&"'"
conn.execute "update 门派 set 结盟='"&co&"' where 门派='"&men&"'"
jmen="<font color=FF0000>【歃血为盟】</font>"&co&"掌门##和"&men&"掌门%%决定结为同盟,共同为一统快乐江湖而努力!"                   
else
jmen="<font color=FF0000>【歃血为盟】</font>"&men&"掌门%%根本不想和你结盟,你要不消灭他们也太丢脸啦!"
end if
else
jmen="<font color=FF0000>【歃血为盟】</font>"&co&"掌门##你到处乱结盟,是不是喝多了!先去和<font color=FF0000>"&ty&"</font>搞清楚再来!"
end if
end if
else
jmen="<font color=FF0000>【歃血为盟】</font>"&co&"的##你不是掌门,是想造反吗!"         
end if
end if
end if
end if
rst.Close
end function
%>
