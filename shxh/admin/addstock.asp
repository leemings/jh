<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�") then Response.Redirect "../error.asp?id=046"
Response.Write "<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false  background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center><font color=0000ff size=4>�����¹�</font></p><hr><form action=addstocknow.asp method=post><table border=3 width=80% align=center><tr><td>��Ʊ����</td><td><input type=text name='stock' size=14 maxlength=7 value=''></td></tr><tr><td>��������</td><td><input type=text value='' size=9 maxlength=9 name='num'></td></tr><tr><td>���м۸�</td><td><input type=text name='price' value='' size=9 maxlength=9></td></tr><tr align=center><td colspan=2><input type=submit name=opt value=' �� �� '> <input type=button onclick=javascript:history.back(); value=' �� �� '></td></tr></table></form></body>"
%>