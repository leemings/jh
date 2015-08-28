<%
function tuoli(un,st)
rst.Open "select 银两,师傅 from 用户 where 姓名='"&un&"'",conn
dejq=rst("银两")
sfa1=rst("师傅")
rst.Close
if sfa1="无" then
tuoli="<font color=FF0000>【离师】</font>##,你并没有师傅，所以你不能进行脱离师徒关系操作！！"
elseif dejq<80000 then
tuoli="<font color=FF0000>【离师】</font>##,现在你没有80万两银子，所以%%不肯答应你与脱离师徒关系！你还是赚够钱再说吧"
elseif sfa1<>st then
tuoli="<font color=FF0000>【离师】</font>%%并不是你的师傅，所以你不能进行脱离师徒关系操作！！"
else
conn.execute "update 用户 set 银两=银两-800000,师傅='无'  where 姓名='"&un&"'"
conn.execute "update 用户 set 银两=银两+800000,徒弟='无'  where 姓名='"&st&"'"
tuoli="<font color=FF0000>【离师】</font>##向原师傅%%交纳了80万块的分手费，终于与%%脱离了师徒关系"
end if
end function
%>
