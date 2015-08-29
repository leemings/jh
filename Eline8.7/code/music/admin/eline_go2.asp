<%PageName="admin"%>
<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<HTML><HEAD><TITLE>后台管理 - 极限视听空间</TITLE>
<meta name=author content=www.06net.com >
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<STYLE>.navPoint {
	COLOR: white; CURSOR: hand; FONT-FAMILY: Webdings; FONT-SIZE: 9pt
}
P {
	FONT-SIZE: 9pt
}
</STYLE>

<SCRIPT>
function switchSysBar(){
	if (switchPoint.innerText==3){
		switchPoint.innerText=4
		document.all("frmTitle").style.display="none"
	}
	else{
		switchPoint.innerText=3
		document.all("frmTitle").style.display=""
	}
}
</SCRIPT>
</HEAD>
<BODY scroll=no style="MARGIN: 0px">
<TABLE border=0 cellPadding=0 cellSpacing=0 height="100%" width="100%">
  <TBODY>
  <TR>
    <TD align=middle id=frmTitle noWrap vAlign=center name="frmTitle"><IFRAME 
     scrolling=auto frameBorder=0 id=BoardTitle name=BoardTitle src="left.asp" 
      style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 150; Z-INDEX: 2"></IFRAME>
    <TD bgColor=#E3005B style="WIDTH: 10pt">
      <TABLE border=0 cellPadding=0 cellSpacing=0 height="100%">
        <TBODY>
        <TR>
          <TD onclick=switchSysBar() style="HEIGHT: 100%"><SPAN class=navPoint 
            id=switchPoint title=关闭/打开左栏>3</SPAN> </TR></TBODY></TABLE></TD>
    <TD style="WIDTH: 100%"><IFRAME frameBorder=0 id=frmright name=frmright 
      scrolling=yes src="welcome.asp" 
      style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1"></IFRAME></TD></TR></TBODY></TABLE></BODY></HTML>

<%
conn.close
set conn=nothing
%></body></html>




