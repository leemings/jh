<%
horsecalc=Request.QueryString("calc")
horsewin=Request.QueryString("win")
bet=Request.QueryString("bet")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
msg="<head><link rel=stylesheet href='../chatroom/css.css'></head><body bgcolor='"&bgcolor&"' background='"&bgimage&"'><table width=50% align=center><tr><td>����ע��"&horsecalc&"</td><td>��ʤ��"&horsewin&"����</td><td>��ע��"&bet&"</td><td>"
win=horsecalc-bet*10
if win>0 then 
	msg=msg&"�䣺"&win
else
	msg=msg&"Ӯ��"&abs(win)
end if	
msg=msg&"</td></tr><tr><td colspan=6 align=center><input type=button onclick=javascript:location.replace('chipin.asp'); value=' �� �� '></td></tr></table></body>"
Response.Write msg
%>