<%
function nenyin(un,st,gr)
if gr<10 then
nenyin="<font color=FF0000>【冷饮】</font>官府的人员对##说：你的等级不到10啊，不能使用这个功能，哈！"
else
rst.Open "SELECT id,体力,内力 FROM 用户 WHERE 姓名='" & un & "' and 会员=true",conn
if rst.EOF or rst.BOF then
nenyin="<font color=FF0000>【冷饮】</font>官府的人员对##说：你不是会员啊，哈！"
else
rst.close
rst.Open "SELECT 等级,体力,体力加,内力加,内力 FROM 用户 WHERE 姓名='" & st & "'",conn
if rst.EOF or rst.BOF then
nenyin="<font color=FF0000>【冷饮】</font>官府的人员对##说：你笨蛋啊,你给谁吃呢，哈！"
elseif rst("体力")>rst("等级")*rst("体力加") then 
nenyin="<font color=FF0000>【冷饮】</font>官府的人员对##说：%%体力超过上限了啊，哈！"
elseif rst("内力")>rst("等级")*rst("内力加") then 
nenyin="<font color=FF0000>【冷饮】</font>官府的人员对##说：%%内力太多,要爆炸了，哈！"
else
nenyin="<font color=FF0000>【冷饮】</font>##对%%说：这位大侠，渴了吧！请吃<IMG src=../chatroom/image/12.gif width=31 height=30>（味道好极了！），呆会再聊！大侠喝了的冷饮体力和内力都增加了2000。"
conn.execute "update 用户 set 体力=体力+2000,内力=内力+2000 where 姓名='" & st & "'"
end if
rst.close
end if
end if
end function
%>
