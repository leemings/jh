<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
chatroomsn=session("yx8_mhjh_userchatroomsn")
chatroomname=Application("yx8_mhjh_systemname"&chatroomsn)
%>
<head>
<link rel=stylesheet href=css.css>
<script language=javascript>
function refresh(){
parent.msgfrm0.document.open();
parent.msgfrm0.document.writeln("<html><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><link rel='stylesheet' href='css.css'></head><\Script Language=JavaScript>var autoScroll=1;function chgautoscroll(){if(!parent.talkfrm.document.talkform.autoscroll.checked){autoScroll=0;}else{autoScroll=1;autoscrollnow();}return true;}function autoscrollnow(){if(autoScroll==1){this.scroll(0,65000);parent.msgfrm1.window.scroll(0,65000);setTimeout('autoscrollnow()',200);}}autoscrollnow();<\/script><body text=000000 oncontextmenu=self.event.returnValue=false><font color=FF0000>【风云变－震天下】</font>欢迎<font color=FF0000>〖<a href='javascript:parent.chgsendto(\"<%=username%>\");' target='talkfrm' onmouseover=\"window.status='选择说话或动作对象';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000><%=username%></font></a>〗</font>光临<%=chatroomname%><font class=timsty><%=time()%></font><br>");parent.msgfrm1.document.open();parent.msgfrm1.document.writeln("<html><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><link rel='stylesheet' href='css3.css'></head><body text=000000 ><font color=FF0000>【风云变－震天下】</font>欢迎<font color=FF0000>〖<a href='javascript:parent.chgsendto(\"<%=username%>\");' target='talkfrm' onmouseover=\"window.status='选择说话或动作对象';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000><%=username%></font></a>〗</font>光临<%=chatroomname%><font class=timsty><%=time()%></font><br>");
parent.getfrm.location.replace('getmsg.asp');
location.replace('onlinelist.asp');
}
</script></head><body onload=refresh();>&nbsp;</body>