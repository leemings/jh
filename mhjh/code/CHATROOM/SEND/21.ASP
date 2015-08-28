<%
function zhutu(un,st)
rst.Open "select 徒弟 from 用户 where 姓名='"&un&"'",conn
sfaw=rst("徒弟")
rst.Close
if sfaw=" " or sfaw="无"  then
zhutu="<font color=FF0000>【逐徒】</font>##你并没有徒弟，逐什么徒啊！！"
elseif sfaw<>st then
zhutu="<font color=FF0000>【逐徒】</font>%%并不是你的徒弟，逐人家干嘛啊！！"
else
conn.execute "update 用户 set 师傅='无'  where 姓名='"&st&"'"
conn.execute "update 用户 set 徒弟='无'  where 姓名='"&un&"'"
zhutu="<font color=FF0000>【逐徒】</font>##把%%逐出了师们，脱离了师徒关系"
end if
end function
%>
