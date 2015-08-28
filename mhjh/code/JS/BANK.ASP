<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
if session("yx8_mhjh_inchat")="" then
Response.write "<script language='javascript'>alert ('你不能进入,请先进入聊天室再来钱庄。谢谢合作');location.href = '';</script>"
Response.End 
end if
chatroomsn=Session("yx8_mhjh_userchatroomsn")
nowdate=cstr(date())
sqlstr="select 银两,存款,结算日期 from 用户 where 姓名='"&username&"'"
set conn=server.CreateObject("adodb.connection")
conn.Mode=16
conn.IsolationLevel=256
conn.Open Application("yx8_mhjh_connstr")
set rst=conn.Execute(sqlstr)
lastdate=rst("结算日期")
if lastdate="" or isnull(lastdate) then lastdate=nowdate
newmoney=clng(rst("存款")*1.01^datediff("d",lastdate,nowdate))
money=rst("银两")
sqlstr="update 用户 set 存款="&newmoney&",结算日期='"&nowdate&"' where 姓名='"&username&"'"
conn.Execute sqlstr
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel="stylesheet" href="../style.css">
<title><%=Application("yx8_mhjh_systemname")%></title>
<script language=javascript>
function check(){
if(document.form1.op[0].checked & parseInt(document.form1.mn.value)<parseInt(document.form1.money.value)){
alert('你的现金不足，无法存入');
return false;
}
if(document.form1.op[1].checked & parseInt(document.form1.nmn.value)<parseInt(document.form1.money.value)){
alert('你的存款不足，无法取出');
return false;
}
return true;
}
</script>
</head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false>
<p align=center><b><font face="方正舒体" size="5" color="#FFFF00"><img border="0" src="../IMAGE/251.GIF"></font><font face="方正舒体" color="#000000" size="4">丰润钱庄</font><font face="方正舒体" size="5" color="#FFFF00"><img border="0" src="../IMAGE/252.gif"></font></b></p>
<form name=form1 method=post action=bankoption.asp onsubmit=return(check())>
<table width='60%' align=center border=1 bordercolor="#990000" cellspacing="1" cellpadding="1" height="132">
<tr><td height="19">现银：</td><td height="19"><input type=text name=mn readonly value='<%=money%>' size=9 class=norstyle></td></tr>
<tr><td height="19">银票：</td><td height="19"><input type=text name=nmn readonly value='<%=newmoney%>' size=9 class=norstyle></td></tr>
<tr><td colspan=2 align=center height="19"><input type=radio name=op checked value='存款'>存款<input type=radio name=op value='取款'>取款</td></tr>
<tr><td colspan=2 height="19" ><input name='money' type=text maxlength=9 size=9></td></tr>
<tr>
<td colspan=2 align=center height="26">
<input type=submit value='执行'></td></tr>
</table>
</form>
<p align=center>
　 
</p> 
<p align="center">　</p> 
</body> 
