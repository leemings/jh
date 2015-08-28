<%
function xingfa(un,st,co,gr,sn,mg)
mg=clng(mg)
if mg>10000 then
xingfa="<font color=FF0000>【处罚】</font>##，你认为对%%处罚,扣掉"&mg&"内力和体力合适么？是不是太多了！！ "
else
if mg<1 then
xingfa="<font color=FF0000>【处罚】</font>##，你认为对%%处罚,扣掉"&mg&"内力和体力合适么？是不是太少了！！ "
else
rst.Open "select 体力 from 用户 where 姓名='"&st&"'",conn
if mg>rst("体力") then
xingfa="<font color=FF0000>【处罚】</font>##，%%的体力连"&mg&"都不到阿，你还要处罚他阿,就可怜可怜,放过%%吧！！ "
elseif co="官府" and gr>=20 then
conn.execute "update 用户 set 体力=体力-"&mg&",内力=内力-"&mg&" where 姓名='"&st&"'"
conn.execute "insert into 聊务记录(管理类型,管理人员,被管理人,管理原因,管理时间,管理房间) values ('刑罚','" & un & "','" & st & "','" & mg & "',now(),'"& Application("yx8_mhjh_systemname"&chatroomsn) &"')"
xingfa="<font color=FF0000>【处罚】</font>官府的##实在忍无可忍%%的行为，对%%处罚"&mg&"的体力和内力，%%大叫阿妈！！ "
else
xingfa="<font color=FF0000>【处罚】</font>你不是我官府的或你级别不够,哈,别耍花样阿!再来就我炸掉你，格式化你的硬盘！ "
end if
end if
end if
end function
%>
