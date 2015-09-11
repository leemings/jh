<%
    wnky=wnk(to1)
if wnky<>"ok" or wnky="zstt" then
 if wnky="zstt" then
  kapian="<font color=green>【卡片】<font color=" & saycolor & ">##想对%%使用["&fn1&"]，但%%已经练成了超凡的能力，此卡片已经被化解了...</font>"
 else
  kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
 end if
conn.execute "update 用户 set w5='"&mycard&"' where  姓名='"&aqjh_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
%>