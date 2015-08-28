<%
function fakuan(un,st,co,gr,sn,mg)
mg=clng(mg)
if mg>500000 then
fakuan="<font color=FF0000>【罚款】</font>##，你认为对%%罚款"&mg&"合适么？太多了！！ "
else
if mg<1 then
fakuan="<font color=FF0000>【罚款】</font>##，你认为对%%罚款"&mg&"合适么？太少了！！"
else
rst.Open "select 银两 from 用户 where 姓名='"&st&"'",conn
if mg>rst("银两") then
fakuan="<font color=FF0000>【罚款】</font>##，%%身上总共只有"&mg&"两，就可怜可怜，放过%%吧！！ "
elseif co="官府" and gr>=20 then
conn.execute "update 用户 set 银两=银两-"&mg&" where 姓名='"&st&"'"
conn.execute "insert into 聊务记录(管理类型,管理人员,被管理人,管理原因,管理时间,管理房间) values ('罚款','" & un & "','" & st & "','" & mg & "',now(),'"& Application("yx8_mhjh_systemname"&chatroomsn) &"')"
fakuan="<font color=FF0000>【罚款】</font>官府的##实在忍无可忍%%的行为，对%%罚款"&mg&"两银子！！ "
else
fakuan="<font color=FF0000>【罚款】</font>你不是我官府的或你级别不够哈! "
end if
end if
end if
end function
%>
