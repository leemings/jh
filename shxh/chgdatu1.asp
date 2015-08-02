<HTML>
<HEAD>
<title><%=Application("Ba_jxqy_systemname")%></title>
<link rel="stylesheet" href="style.css">
</HEAD>
<body  bgcolor="<%=Application("Ba_jxqy_backgroundcolor")%>" oncontextmenu=self.event.returnValue=false text="#000000" background="<%=Application("Ba_jxqy_backgroundimage")%>">
<p align=center> <font color=#ff0000 size=4>更改资料</font></p>
<form action="getdatum.asp" method=post id=form1 name=form1>
<table width="180" bgcolor="#cccccc" align=center border=3>
<tr><td width="60">帐号</td><td width="120"><input name=account maxlength=14 size=14 class="norsty" 
     style="FONT-FAMILY: sans-serif" 
   ></td></tr>
<tr><td>姓名</td><td width="118"><input name=username maxlength=14 size=14 class="norsty" 
     ></td></tr>
<tr><td>密码</td><td width="118"><input type=password name=password maxlength=14 size=14 class="norsty"></td></tr>
<tr><td align=middle colspan=2 ><input type=submit value=" 提 交 " class="norsty" id=submit1 name=submit1> <INPUT type=button onclick=javascript:top.window.close(); value=" 关 闭 " class="norsty" id=button1 name=button1></td></tr> 
</table> 
</form> 
</body></HTML> 
