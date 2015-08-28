<%
function fchwg(un,st,co)
rst.Open "select id from 用户 where 姓名='"&un&"' and 门派='"&co&"' and 身份='掌门'",conn
if rst.EOF or rst.BOF then
fchwg="<font color=FF0000>【废除武功】</font>##,你不是本派掌门，无法废除%%的武功"
else
rst.Close
rst.Open "select * from 用户 where 姓名='"&st&"' and 门派='"&co&"' and 身份<>'掌门'",conn
if rst.EOF or rst.BOF then
fchwg="<font color=FF0000>【废除武功】</font>%%不是本派弟子，所以##你无权废除%%的武功！"
else
conn.execute "delete  from 攻击  where  姓名='"&st&"'"
fchwg="<font color=FF0000>【废除武功】</font>由于%%犯了严重的错误，##废除%%的武功！"
end if
end if
rst.Close
end function
%>
