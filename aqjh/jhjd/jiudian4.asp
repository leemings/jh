<%@ LANGUAGE=VBScript codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
id=request("id")%>
<HTML>
<HEAD>
<TITLE>请输入你要宴请朋友的名字</TITLE>
<script language=JavaScript>
document.onmousedown=click;
function click(){if(event.button==2){if(confirm("是否关闭窗口？","江湖提示")){window.close();}}}
</script>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK
href="regimage/home.css" type=text/css rel=stylesheet>
<STYLE>TD {
FONT-SIZE: 12px; LINE-HEIGHT: 16px
}
input {border: 1px solid; font-size: 9pt; font-family: "宋体"; border-color:
#000000 solid
}
</STYLE>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
</HEAD>
<BODY text=#cccccc vLink=#ffff00 aLink=#ffff00 link=#ffff00 bgColor=#3a4b91
leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<TABLE cellSpacing=0 cellPadding=0 width=400 align=left border=0>
<TBODY>
<TR>
<TD vAlign=top height=365>
<TABLE cellSpacing=0 cellPadding=0 width=400
background=regimage/top12.gif border=0>
<TBODY>
<TR>
<TD vAlign=top height=105>
<P>&nbsp;</P>
<P>　　　　　说明：请输入你要宴请的朋友的名字。 </P></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=360 align=center border=0>
<FORM action="jiudian44.asp" method=post>
<TBODY>
<TR>
<TD class=p9> 姓名：
<INPUT size=10 name=name maxlength="10">
<INPUT size=2 readonly type="hidden" name=id maxlength="2" value=<%=id%>>
</TD></TR>
<TR>
<TD class=p11 height=1></TD></TR>
<TR>
<TD
class=p9>
</TD></TR>
<TR>
<TD class=p11 height=1></TD></TR>
<TR>
<TD class=p9 align=middle><INPUT type=submit value=确认> <INPUT onclick=window.close() type=button value=关闭>
</TD></TR></FORM></TBODY></TABLE>
<P></P>
<TABLE cellSpacing=0 cellPadding=0 width=400 border=0>
<TBODY>
<TR>
<TD></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></BODY></HTML>