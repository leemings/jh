<%
function distributemoney(un,mg)
rst.Open "select 国家,职位 from 用户 where 姓名='"&un&"'",conn
guojia=rst("国家")
zhiwei=rst("职位")
if zhiwei<>"君主" then
distributemoney="<font color=FF0000>【君主赈灾】</font>你不是君主,你无权赈灾"
else
rst.Close
rst.Open "select 管理时间 from 聊务记录",conn
sj=rst("管理时间")
nosaytime=datediff("s",sj,now())
if nosaytime<1500 then
s=1500-nosaytime
 distributemoney="<font color=FF0000>【君主赈灾】</font>刚刚有国家赈灾完毕,你的国家无法赈灾,目前离下次赈灾时间还有<font color=FF0000>"&s&"</font>秒!"
else
randomize()
money=clng(rnd()*2000)+1000000
onlinelist=Application("yx8_mhjh_onlinelist")
for i=1 to ubound(onlinelist) step 6
conn.execute "update 用户 set 银两=银两+"&money&" where 姓名='"&onlinelist(i)&"' and 国家='"&guojia&"'"
next
erase onlinelist
conn.execute "update 聊务记录 set 管理类型='发钱',管理人员='" & un & "',被管理人='本国国民',管理原因='"&money&"',管理时间=now(),管理房间='"& Application("yx8_mhjh_systemname"&chatroomsn) &"'"
distributemoney="<table align=center><td><table border=0 cellpadding=0 cellspacing=0 ><tr><td><img src=../chatroom/image/rightct3.gif ></td><td background=../chatroom/image/rightct4.gif ></td> <td><img  src=../chatroom/image/rightct1.gif ></td></tr><tr><td background=../chatroom/image/rightct8.gif ></td><td valign=center align=center><table align=center border=0 cellpadding=1 cellspacing=0 ><tr> <td valign=center align=center><font color=red>【君主赈灾】</font></td> </tr> <tr><td valign=center align=center >"&guojia&"国君主##开仓赈灾,在线"&guojia&"国国民每人得到银两<bgsound src='../mid/faqian.wav' loop=1>"&money&"</td></tr> <tr><td valign=center align=center ></td></tr></table></td><td background=../chatroom/image/rightct08.gif></td></tr><tr><td><img src=../chatroom/image/rightct9.gif></td><td background=../chatroom/image/rightct10.gif></td><td><img src=../chatroom/image/rightct11.gif></td></tr></table></td></table></body>"
end if
end if
rst.Close
end function
%>
