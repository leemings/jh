<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
msg="<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel='stylesheet' href='style.css'></head><body bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu=self.event.returnValue=false><p align=center><font color=ff0000 size=5>断 魂 涯</font></p>"
degree=Request.QueryString("degree")
if degree="" then 
	degree=-1
elseif not isnumeric(degree) then
	degree=-1
end if
degree=clng(degree)+1
suicide=array("我想自杀？","我的确不想继续我无谓的生命了","我下定决心--自杀是个很好的选择")
abandontxt=array("关 闭","生命很可贵","想想还是算了")
if degree>1 then 
	msg=msg&"<form method=post action=suicide2.asp>"
else
	msg=msg&"<form method=post action=suicide.asp?degree="&degree&">"
end if
msg=msg&"<table border=0  width=80% align=center><tr align=center><td><img src='images/death_head.gif'></td></tr><tr align=center><td><input type=submit value=' "&suicide(degree)&" '> <input type='button' value=' "&abandontxt(degree)&" ' onclick='javascript:top.window.close();' ></td></tr></table></form></body>"
Response.Write msg
%>