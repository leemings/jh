<%
function degree(un,st,co,mg)
if mg="掌门" then
degree="<font color=FF0000>【册封】</font>##,你无法册封%%为掌门"
else
rst.Open "select id from 用户 where 姓名='"&un&"' and 门派='"&co&"' and 身份='掌门'",conn
if rst.EOF or rst.BOF then
degree="<font color=FF0000>【册封】</font>##,你不是本派掌门，无法册封%%"
else
rst.Close
rst.Open "select id from 用户 where 姓名='"&st&"' and 门派='"&co&"' and 身份<>'掌门'",conn
if rst.EOF or rst.BOF then
degree="<font color=FF0000>【册封】</font>%%不是本派弟子，所以##无法册封！"
else
conn.execute "update 用户 set 身份='"&mg&"' where  姓名='"&st&"'"
degree="<font color=FF0000>【册封】</font>##册封%%为"&co&mg&"！"
end if
end if
rst.Close
end if
end function
%>
