<%
    wnky=wnk(to1)
if wnky<>"ok" or wnky="zstt" then
 if wnky="zstt" then
  kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##���%%ʹ��["&fn1&"]����%%�Ѿ������˳������������˿�Ƭ�Ѿ���������...</font>"
 else
  kapian="<font color=green>����Ƭ��<font color=" & saycolor & ">##͵͵��%%ʹ����["&fn1&"]...</font>"&wnky
 end if
conn.execute "update �û� set w5='"&mycard&"' where  ����='"&aqjh_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','����','"& fn1 & "')"
exit function
end if
%>