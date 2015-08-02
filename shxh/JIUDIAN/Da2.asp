<!--#include file="dadata.asp"-->
<%
nickname=session("Ba_jxqy_username")
sql="select * from 用户 where 姓名='" & nickname & "'"
set rs=conn.execute(sql)
if rs.bof or rs.eof then
	response.write "你不是剑侠中人，不能进参加宴会"
	conn.close
	response.end
else
dj=rs("等级")
if dj<2 then
	response.write "你还是江湖小辈，不能参加宴会"
	conn.close
	response.end
else

id=request("id")
sql="select * from 宴会列表 where ID=" & id
Set Rs=connt.Execute(sql)
yyou=rs("拥有者")
nl=rs("内力")
tl=rs("体力")
jb=rs("金币")
lx=rs("类型")
ls=rs("数量")
if ls<1 then 
sql1="delete * from 宴会列表  where ID=" & id
connt.execute sql1
sql1="delete * from 宴会者 where 宴会名=" & id
connt.execute sql1
response.write "你来晚，宴会已经结束"
	connt.close
	response.end
else

sql="select * from 宴会者 where 参加者='" & nickname & "'  and 宴会名=" & id 
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
sql2="insert into 宴会者(参加者,宴会名) values ('" & nickname & "' , " & id & " )"
	connt.execute sql2
sql1="update 用户 set 内力=内力+"&nl&",体力=体力+"&tl&",精力=精力+"&jb&" where 姓名='"&nickname&"' "
conn.execute sql1
sql1="update 宴会列表 set 数量=数量-1 where ID=" & id
connt.execute sql1
conn.close
connt.close
set rs=nothing
Application.UnLock
talkarr=Application("Ba_jxqy_talkarr")
talkpoint=Application("Ba_jxqy_talkpoint")
dim newtalkarr(600)
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
newtalkarr(594)=username
newtalkarr(595)="大家"
newtalkarr(596)=""
newtalkarr(597)="000000"
newtalkarr(598)="000000"
newtalkarr(599)="<font color=FF0000>【公告】</font>"&msg&"<font class=timsty>"&nowtime&"</font>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("Ba_jxqy_talkpoint")=talkpoint+1
Application("Ba_jxqy_talkarr")=newtalkarr
Application.UnLock
%>

<html>

<head>
<title>参加宴会成功</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
.unnamed1 {  font-size: 9pt}
-->
</style>

</head>

<body bgcolor="#CCCCCC">
<p>&nbsp;</p>
<table width="90%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="17" bgcolor="#996633" align="center"><font color="#FFFFFF">参加酒宴成功</font></td>
  </tr>
  <tr bgcolor="#66FF66"> 
    <td align="center" height="157"> 
      <p><font> <img src="jd.files/ka1.gif" width="204" height="251"></font></p>
    </td>
  </tr>
  <tr bgcolor="#0033CC"> 
    <td align="center" height="20" class="unnamed1"><b><font color="#FF3333">你经过了一番狼吞虎咽，喝的上面一个模样，增加内力，体力，精力都涨了，自己看看涨了多少吧！呵呵"</font></b></td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>

<%else%>
<script>
alert('你已参加过这个宴会了，怎么还来啊。');
window.close();
</script>
<%end if%>
<%end if%>
<%end if%>
<%end if%>