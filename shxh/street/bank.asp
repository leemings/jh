<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
nowdate=cstr(date())
sqlstr="select 银两,存款,结算日期 from 用户 where 姓名='"&username&"'"
set conn=server.CreateObject("adodb.connection")
conn.Mode=16
conn.IsolationLevel=256
conn.Open Application("Ba_jxqy_connstr")
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
<title><%=Application("Ba_jxqy_systemname")%></title>
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
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color=0000FF size=5><%=Application("Ba_jxqy_systemname")%></font></p>
<form name=form1 method=post action=bankoption.asp onsubmit=return(check())>
    <table width='60%' align=center border=1 bordercolor="#990000" cellspacing="1" cellpadding="1">
      <tr><td>现银</td><td><input type=text name=mn readonly value='<%=money%>' size=9 class=norstyle></td></tr>
<tr><td>银票</td><td><input type=text name=nmn readonly value='<%=newmoney%>' size=9 class=norstyle></td></tr>
<tr><td colspan=2 align=center><input type=radio name=op checked value='存款'>存款<input type=radio name=op value='取款'>取款</td></tr>
<tr><td colspan=2 ><input name='money' type=text maxlength=9 size=9></td></tr>
<tr>
        <td colspan=2 align=center height="32">
<input type=submit value='执行'></td></tr>
</table>
</form>
</div>
<p align=center>
  <input type="button" value=" 返 回 " onClick="javascript:location.href='street.asp'" name="button"> 
  <input type="button" value=" 关 闭 " onClick="javascript:window.close();" name="button">
</p>
</body>