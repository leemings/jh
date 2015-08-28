<%
function dislodge(un,st,co)
rst.Open "select id from 用户 where 姓名='"&un&"' and 门派='"&co&"' and 身份='掌门'",conn
if rst.EOF or rst.BOF then
dislodge="<font color=FF0000>【逐出】</font>##,你不是本派掌门，无法将%%逐出山门"
else
rst.Close
rst.Open "select id from 用户 where 姓名='"&st&"' and 门派='"&co&"' and 身份<>'掌门'",conn
if rst.EOF or rst.BOF then
dislodge="<font color=FF0000>【逐出】</font>%%不是本派弟子，所以##你无权处置！"
else
conn.execute "update 用户 set 身份='无',门派='无' where  姓名='"&st&"'"
dislodge="<font color=FF0000>【逐出】</font>##将%%无情地逐出了"&co&"！"
end if
end if
rst.Close
end function
%>
