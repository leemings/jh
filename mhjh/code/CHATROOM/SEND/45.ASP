<%
function duanyi(un,st)
rst.Open "select 结拜兄弟,银两 from 用户 where 姓名='"&un&"'",conn
jbxd=rst("结拜兄弟")
yl=rst("银两")
rst.Close
if jbxd="无" then
duanyi="<font color=FF0000>【割袍断义】</font>##,你并没有结拜兄弟，所以你不能进行割袍断义！"
elseif yl<80000 then
duanyi="<font color=FF0000>【割袍断义】</font>##,现在你没有8万两银子，所以%%不肯答应与你断绝兄弟关系！你还是赚够钱再说吧"
elseif jbxd<>st then
duanyi="<font color=FF0000>【割袍断义】</font>%%并不是你的结拜兄弟，所以你不能进行脱割袍断义！"
else
conn.execute "update 用户 set 银两=银两-70000,结拜兄弟='无'  where 姓名='"&un&"'"
conn.execute "update 用户 set 银两=银两+70000,结拜兄弟='无'  where 姓名='"&st&"'"
duanyi="<font color=FF0000>【割袍断义】</font>##向兄弟%%交纳了7万块的分手费，终于与%%脱离了兄弟关系"
end if
end function
%>
