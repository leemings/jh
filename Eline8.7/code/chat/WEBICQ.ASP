<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
sjjh_name=Session("sjjh_name")
sjjh_jhdj=Session("sjjh_jhdj")
inroom=session("inroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Session("sjjh_inthechat")<>"1" then
Response.Write "<script Language=Javascript>alert('进入聊天室再呼叫！');window.close();</script>"
Response.End
end if
if sjjh_jhdj<10 then
Response.Write "<script Language=Javascript>alert('战斗等级不够10级不可呼叫！');window.close();</script>"
Response.End
end if
chatroombgimage=Application("sjjh_chatimage")
chatroombgcolor=Application("sjjh_chatbgcolor")
if Request.Form("towho")<>"" then
towho=trim(Request.Form("towho"))
intro=trim(Request.Form("intro"))
url=lcase(trim(Request.Form("url")))
checksong=trim(Request.Form("checksong"))
checktv=trim(Request.Form("checktv"))
kuan=clng(Request.Form("kuan"))
gao=clng(Request.Form("gao"))
intro=replace(intro,chr(39),"")
intro=replace(intro,chr(34),"")
if checksong="" then checksong="off"
if checktv="" then checktv="off"
'Response.Write url
'Response.End
saystr="parent.webicq('" & sjjh_name & "','" & towho & "','" & intro & "','" & now() & "','" & url & "','" & checksong & "','" & checktv & "'," & kuan & "," & gao & ");"
act="消息"
towhoway=0
addwordcolor="660099"
saycolor="008888"
addsays="对"
says="<font class=t2>【呼叫网友】</font>##向%%发送了呼叫消息……"
zj="<a href=javascript:parent.sw(\'[" & sjjh_name & "]\'); target=talk><font color=" & addwordcolor & ">" & sjjh_name & "</font></a>"
br="<a href=javascript:parent.sw(\'[" & towho & "]\'); target=talk><font color=" & addwordcolor & ">" & towho & "</font></a>"
says=Replace(says,"##",zj)
says=Replace(says,"%%",br)
saystr=saystr&"parent.sh('" & addwordcolor & "','" & saycolor & "','" & act & "','" & sjjh_name & "','" & addsays & "','" & towho & "','" & says & "','" & towhoway & "');"
Application.Lock
sd=Application("sjjh_sd" & inroom)
line=int(Application("sjjh_line" & inroom))
Application("sjjh_line" & inroom)=line+1
for i=1 to 48
sd(i)=sd(i+2)
next
sd(49)=line+1
sd(50)=saystr
Application("sjjh_sd" & inroom)=sd
Application.UnLock
Response.Write "<script language=JavaScript>{alert('消息发送成功！');window.close()}</script>"
Response.End
end if
%><html>
<head>
<title>呼叫聊友</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
<script Language=JavaScript>
function check(){
if(document.forms[0].intro.value==""){alert("消息内容不能为空！");return false;}
if(document.forms[0].gao.value==""){alert("高不能为空！");return false;}
if(document.forms[0].kuan.value==""){alert("宽不能为空！");return false;}
var pattern = /^([0-9])+$/;
if(pattern.test(document.forms[0].gao.value)!=true){alert("提示：高请用数字！");return false;}
if(pattern.test(document.forms[0].kuan.value)!=true){alert("提示：高请用数字！");return false;}
document.forms[0].Submit.disabled=1;return true;
}
function textCounter(field, countfield, maxlimit) {
if (field.value.length > maxlimit)field.value = field.value.substring(0, maxlimit);
else countfield.value = maxlimit - field.value.length;}
function DoTitle(addTitle) {
document.form.url.value=addTitle;
document.form.url.focus();
return; }
</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=<%=chatroombgcolor%> background=<%=chatroombgimage%> leftmargin="0" topmargin="3">
<table width=100% border=1 align=center cellspacing=0 bordercolorlight=000000 bordercolordark=FFFFFF bgcolor=E0E0E0>
<tr valign="top">
<td> <table border=0 bgcolor=#009933 cellspacing=0 cellpadding=2 width=100%>
<tr>
<td width=326><font color=FFFFFF face=Wingdings>*</font><font color=FFFFFF>发送
Web ICQ 消息</font></td>
<td width=18> <table border=1 bordercolorlight=666666 bordercolordark=FFFFFF cellpadding=0 bgcolor=E0E0E0 cellspacing=0 width=18>
<tr>
<td width=16><b><a href="javascript:window.close()" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="退出"><font color="000000">×</font></a></b></td>
</tr>
</table></td>
</tr>
</table>
<table width="100%" border="1" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" bgcolor="FFFFFF" cellpadding="0">
<tr valign="middle" align="center">
<td class=p9> <table width="100%" border="0" cellpadding="3">
<tr align="center">
<td width="25%" bgcolor="#009900"><font color="FFFFFF">发送消息</font></td>
<td>&nbsp;</td>
</tr>
</table>
<table width="100%" border="1" height="200" cellspacing="0" cellpadding="5" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="E0E0E0">
<form method="post" action="webicq.asp" onsubmit='return(check());' name="form">
<tr>
<td> <div align="center"><font color="#993300">限制输入1024字，现在您还可输入:
<input readonly type=text name=remLen size=4 maxlength=3 value="1024" style="BORDER-BOTTOM-WIDTH: 0px; BORDER-LEFT-WIDTH: 0px; BORDER-RIGHT-WIDTH: 0px; BORDER-TOP-WIDTH: 0px; COLOR: red" >
字节 </font></div>
<table width="100%" border="0" align="center">
<tr>
<td>对象:</td>
<td> <select name="towho">
<%useronlinename=Application("sjjh_useronlinename" & inroom)
show=Split(Trim(useronlinename),",")
x=UBound(show)
for i=0 to x-1%>
<option value="<%=show(i)%>"><%=show(i)%></option>
<%next%></select>共<font color=red><%=x+1%></font>人 <font color="#0000FF">音乐:</font><input name="checksong" type="checkbox"><font color="#0000FF">影视:</font><input name="checktv" type="checkbox">
</td>
</tr>
<tr>
<td>网址:</td>
<td><input name="url" type="text" id="url" value="http://" size="40" maxlength="150"></td></tr>
<tr><td>列表:</td>
<td>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
%>
<select name="select" onchange=DoTitle(this.options[this.selectedIndex].value)>
<option value="http://" selected>快速操作</option>
<option value="" selected style='color:blue'>〖音乐类〗</option>
<%sql="SELECT * FROM ss where c='音乐' order by d"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
<option value="<%=rs("b")%>"><%=rs("a")%></option>
<%rs.movenext
loop
rs.close%>
<option value="" style='color:blue' >〖falsh动画〗</option>
<%sql="SELECT * FROM ss where c='flash' order by d"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
<option value="<%=rs("b")%>"><%=rs("a")%></option>
<%rs.movenext
loop
rs.close%>
<option value="" style='color:blue' >〖图片贴图〗</option>
<%sql="SELECT * FROM ss where c='贴图' order by d"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
<option value="<%=rs("b")%>"><%=rs("a")%></option>
<%rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing%>
	</select>影像宽:<input name="kuan" type="text" value="400" size="3" maxlength="3">高:<input name="gao" type="text" size="3" value="340" maxlength="3"></td>
</tr><tr>
<td>消息:</td>
<td> <textarea name="intro" cols="40" rows="7"  onKeyDown="textCounter(this.form.intro,this.form.remLen,1024);" onKeyUp="textCounter(this.form.intro,this.form.remLen,1024);"></textarea>
</td>
</tr>
<tr>
<td colspan="2" align="center"> <input type="submit" name="Submit" value="发送" accesskey='s'>
<input name="reset" type="reset" value="重写"> </td>
</tr>
</table></td>
</tr>
</form>
</table></td>
</tr>
</table></td>
</tr>
</table>
</body>
</html>