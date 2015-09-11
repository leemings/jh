<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 金币 from [用户] where 姓名='"& aqjh_name &"'",conn,1,1
jhjb=rs("金币")
rs.close
set rs=nothing
conn.close
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from use where a='"& aqjh_name &"'", conn, 1,3
showjb=rs("e")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<link rel="stylesheet" href="../css.css">
<title>江湖秀账号充值</title>
<SCRIPT language=JavaScript>
function check(){
var money=document.form.money.value;
var pattern = /^([0-9])+$/;
if(pattern.test(money)!=true){alert("提示：金币数使用数字！");return false;}
document.form.bb.disabled=1;
}
</script>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bgcheetah.gif">
<div align="center"><font color="#000000"><b>欢迎<font color=blue><%=aqjh_name%></font>光临江湖秀账号充值</b></font></div>
<br>
<table width="300" border="1" align="center" bordercolor="#000000">
  <tr>
<td valign="top" height="8" >
<form method="POST" action="hyok.asp" name="form" onSubmit="return check(this);">
<div align="center">
<p>现有金币：<%=jhjb%>个，江湖秀币:<%=showjb%>个<br>
<br>
<input type="text" name="money" size="10" maxlength="5">
<input type="submit" value="转换" name="bb" class="p9">
</p>
</div>
</form>
</td>
</tr>
<tr>
<td valign="top" >
<p>江湖金币充值转换江湖秀1/10，即1个金币可以转换10个江湖秀币，转换后的秀币不可以转换成江湖金币。</p>
</td>
</tr>
</table>
<div align="center"><font color="#00FF66"><b><font color="#0000FF"><%=Application("aqjh_chatroomname")%></font></b></font>
</div>
</body>
</html>