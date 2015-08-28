<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=session("yx8_mhjh_username")
if name="" then Response.Redirect "../../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 物品 WHERE 所有者='" & name & "' and 数量>=50 and 名称='冰水'  "
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then %>
<script language=javascript>setTimeout('history.back();',10);</script>
<script language="vbscript">alert("您冰水不够50个，所以不能操作！")
</script><%
response.end
end if
sql="SELECT * FROM 物品 WHERE 所有者='" & name & "' and 数量>=50 and 名称='铁块'  "
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then %>
<script language=javascript>setTimeout('history.back();',10);</script>
<script language="vbscript">alert("您铁块不够50个，所以不能操作！")
</script><%
response.end
end if
sql="SELECT * FROM 物品 WHERE 所有者='" & name & "' and 数量>=20 and 名称='屠龙刀'  "
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then %>
<script language=javascript>setTimeout('history.back();',10);</script>
<script language="vbscript">alert("您屠龙刀不够20个，所以不能操作！")
</script><%
response.end
end if
sql="SELECT * FROM 物品 WHERE 所有者='" & name & "' and 数量>=10 and 名称='小鲤鱼'  "
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then %>
<script language=javascript>setTimeout('history.back();',10);</script>
<script language="vbscript">alert("您小鲤鱼不够10条，所以不能操作！")
</script><%
response.end
end if
sql="update 物品 set 数量=数量-50 WHERE 所有者='" & name & "' and 名称='冰水'  "
Set Rs=conn.Execute(sql)
sql="update 物品 set 数量=数量-50 WHERE 所有者='" & name & "' and 名称='铁块'  "
Set Rs=conn.Execute(sql)
sql="update 物品 set 数量=数量-20 WHERE 所有者='" & name & "' and 名称='屠龙刀'  "
Set Rs=conn.Execute(sql)
sql="update 物品 set 数量=数量-10 WHERE 所有者='" & name & "' and 名称='小鲤鱼'  "
Set Rs=conn.Execute(sql)
sql="select * from 物品 where 名称='天王之刃' and 所有者='" & name & "' and 属性='武器'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof then
sql="insert into 物品(名称,所有者,数量,属性,特效,体力,价格,防御,攻击,内力,寄售,装备) values ('天王之刃','" & name &"',1,'武器','无',0,200000000,0,2000000,0,False,False)"
conn.execute(sql)
	kl="<font color=blue>功夫不负苦心人，"&name&"的<font color=red>天王之刃</font>修练完成！装备上能增加攻击200万！看来江湖中又多了一位高手,血雨腥风的日子又要来临了!</font> ！"
else
sql="update 物品 set 数量=数量+1 where 名称='天王之刃' and 所有者='"&name&"'"
conn.execute(sql)
	kl="<font color=blue>功夫不负苦心人，"&name&"的<font color=red>天王之刃</font>修练完成！装备上能增加攻击200万！看来江湖中又多了一位高手,血雨腥风的日子又要来临了!</font> ！"
end if
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>【铸造武器】</b></font>"&kl	
dim newtalkarr(600) 
talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=name
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="008000" 
		newtalkarr(599)=""&says&""
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
%>
<script language=javascript>setTimeout('history.back();',10);</script>
<script Language="Javascript">alert("您天王之刃修练完成！")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>
