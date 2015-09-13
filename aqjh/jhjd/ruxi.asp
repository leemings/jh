<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('刷钱是吗？喜欢是吗？点啊，刷啊！！');i=i+1;}top.location.href='../exit.asp'}</script>
<!--#include file="../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<!--#include file="dadata.asp"-->
<%
sql="select * from 用户 where 姓名='"&aqjh_name&"'"
set rs=conn.execute(sql)
if rs.bof or rs.eof then
response.write "你不是江湖中人，不能进参加宴会"
conn.close
response.end
else
mypai=rs("门派")
dj=rs("等级")
if dj<2 then
conn.close
Response.Write "<script language=javascript>alert('你还是江湖小辈，不能参加宴会');window.close();</script>"	
response.end
else
id=request("id")
sql="select * from 宴会列表 where ID=" & id & " and 门派='所有' and 参加者='所有'"
Set Rs=connt.Execute(sql)
if rs.bof or rs.eof then
connt.close
Response.Write "<script language=javascript>alert('本席不是为你定的，不要乱来！这样不好玩:P');window.close();</script>"	
response.end
else
yhming=rs("宴会名")
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
connt.close
Response.Write "<script language=javascript>alert('你来晚，宴会已经结束!');history.back();</script>"	
response.end
else
sql="select * from 宴会者 where 参加者='" & aqjh_name & "'  and 宴会名=" & id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
sql2="insert into 宴会者(参加者,宴会名) values ('" & aqjh_name & "' , " & id & " )"
connt.execute sql2
sql1="update 用户 set 内力=内力+"&nl&",体力=体力+"&tl&" where 姓名='"&aqjh_name&"' "
conn.execute sql1
sql1="update 宴会列表 set 数量=数量-1 where ID=" & id
connt.execute sql1
conn.close
connt.close
set rs=nothing
says="<font color=red>【小道消息】</font><font color=green>"&aqjh_name&"参加了"&yyou&"在江湖大酒店"&yhming&"厅举行的※"&lx&"※宴会，增长了不少的体力和内力。</font>"			'聊天数据
call showchat(says)
else
connt.close
set connt=nothing
Response.Write "<script language=javascript>alert('你已参加过这个宴会了，怎么还来啊!');history.back();</script>"
response.end
end if%>
<%end if%>
<%end if%>
<%end if%>
<%end if%>
<html>
<title>参加宴会成功...欢迎光临快乐江湖 http://www.happyjh.com</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
.unnamed1 {  font-size: 9pt}
input {border: 1px solid; font-size: 9pt; font-family: "宋体"; border-color: #000000 solid}
-->
</style>
<body bgcolor="#CCCCCC">
<p>&nbsp;</p>
<table width="90%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
<tr>
<td height="17" bgcolor="#996633" align="center"><font color="#FFFFFF">参加酒宴成功</font></td>
</tr>
<tr bgcolor="#66FF66">
<td align="center" height="157">
<p><font> <img src="jd/ka1.gif" width="204" height="251"></font></p>
</td>
</tr>
<tr bgcolor="#0033CC">
<td align="center" height="20" class="unnamed1"><b><font color="#FF3333">你经过了一番狼吞虎咽，喝的上面一个模样，增加内力 <%=nl%>，体力 <%=tl%>,领到礼金？？？个金币</font></b></td>
</tr>
<tr bgcolor="#0033CC">
<td align="center" height="20"><b><font color="#FF3333"><input  onClick="javascript:window.document.location.href='jd.asp'" value="返 回" type=button name="back">  <INPUT onclick=window.close() type=button value=关闭>  </font></b></td>
</tr>
</table>
<p>&nbsp;</p>
</body>
</html>