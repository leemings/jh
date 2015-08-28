<%
function chuanwei(un,st,co)
rst.Open "select 等级 from 用户 where 姓名='"&un&"' and 门派='"&co&"' and 身份='掌门' and 银两>10000000",conn
if rst.EOF or rst.BOF then
chuanwei="<font color=FF0000>【传位】</font>##,你不是本派掌门或者银两不过1百万，无法传位给%%"
else
uhp=c
hp=rst("等级")
rst.Close
rst.Open "select 等级,体力,银两 from 用户 where 姓名='"&st&"' and 门派='"&co&"' and 身份<>'掌门'",conn
if rst.EOF or rst.BOF then
chuanwei="<font color=FF0000>【传位】</font>%%不是本派弟子，所以##无法传位！"
else if hp>rst("等级") or uhp>rst("体力") then
chuanwei="<font color=FF0000>【传位】</font>%%不够##厉害，##不能传位"
else if rst("银两")<10000000 then
chuanwei="<font color=FF0000>【传位】</font>%%银两不够，双方都需要交纳一亿税务"
else
conn.execute "update 用户 set 银两=银两-1000000,身份='掌门' where  姓名='"&st&"'"
conn.execute "update 用户 set 银两=银两-1000000,身份='元老' where  姓名='"&un&"'"
chuanwei="<font color=FF0000>【传位】</font>##把"&co&"的掌门之位传给了%%！自己变为元老了，双方银两各减少1百万1"
end if
end if
end if
end if
rst.Close
end function
%>
