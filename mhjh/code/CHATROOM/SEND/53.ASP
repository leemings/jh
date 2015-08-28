<%
function fengwei(un,st,mg)
if mg="君主" then
fengwei="<font color=FF0000>【封位】</font>##,你无法册封%%为君主"
else
rst.Open "select 国家 from 用户 where 姓名='"&un&"' and 职位='君主'",conn
if rst.EOF or rst.BOF then
fengwei="<font color=FF0000>【封位】</font>##,你不是君主，无法册封%%的职位"
else
gj=rst("国家")
rst.Close
rst.Open "select id from 用户 where 姓名='"&st&"' and 国家='"&gj&"' and 职位<>'君主'",conn
if rst.EOF or rst.BOF then
fengwei="<font color=FF0000>【封位】</font>%%不是本国子民，所以##无法册封！"
else
conn.execute "update 用户 set 职位='"&mg&"' where  姓名='"&st&"'"
fengwei="<font color=FF0000>【封位】</font>##册封%%为"&gj&mg&"！"
end if
end if
rst.Close
end if
end function
%>
