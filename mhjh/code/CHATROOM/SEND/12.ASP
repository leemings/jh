<%
function exaltgrade(un,st,co,gr,sn)
if co="官府" and gr>=Application("yx8_mhjh_exaltgraderight") then
rst.open "select 等级 from 用户 where 姓名='"&st&"'",conn
if rst("等级")<5 or gr=Application("yx8_mhjh_allright") then
conn.Execute "update 用户 set 等级=等级+1 where 姓名='"&st&"'"
tt=now()
lt=dateadd("n",10,tt)
aberrantlist=Application("yx8_mhjh_aberrantlist")
aberrantlistubd=ubound(aberrantlist)
dim newaberrantlist()
newaberrantname=";"
j=1
for i=1 to aberrantlistubd step 4
if datediff("s",aberrantlist(i+3),tt)<0 then
redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
newaberrantlist(j)=aberrantlist(i)
newaberrantlist(j+1)=aberrantlist(i+1)
newaberrantlist(j+2)=aberrantlist(i+2)
newaberrantlist(j+3)=aberrantlist(i+3)
newaberrantname=newaberrantname&aberrantlist(i)&";"
j=j+4
end if
next
redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
newaberrantlist(j)=st
newaberrantlist(j+1)=un
newaberrantlist(j+2)="升级"
newaberrantlist(j+3)=lt
newaberrantname=newaberrantname&st&";"
Application.Lock
Application("yx8_mhjh_aberrantlist")=newaberrantlist
Application("yx8_mhjh_aberrantname")=newaberrantname
Application.UnLock
erase  newaberrantlist
erase aberrantlist
exaltgrade="<font color=FF0000>【升级】</font>##成功的将%%的等级提升了一级。"
conn.execute "insert into 聊务记录(管理类型,管理人员,被管理人,管理原因,管理时间,管理房间) values ('升级','" & un & "','" & st & "','升级',now(),'"& Application("yx8_mhjh_systemname"&chatroomsn) &"')"
else
exaltgrade="<font color=FF0000>【升级】</font>官府主管笑哈哈地对##说：'%%的等级已经达到或超过了五级，无法再行提升。'"
end if
rst.Close
else
exaltgrade="<font color=FF0000>【升级】</font>##的等级不够，无法替%%升级。"
end if
end function
%>
