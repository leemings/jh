<%@ LANGUAGE=VBScript codepage ="936" %>
<%

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"


sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"

my=sjjh_name
%><!--#include file="dadata.asp"-->
<%
sql="select * from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.write "你不是江湖中人，不能订购酒宴"
conn.close
set conn=nothing
response.end
else
id=request("id")
sex=rs("性别")
if sex="男" then
sex="女"
else
sex="男"	
end if	
%>
<HTML>
<HEAD>
<TITLE>请输入你要宴请的<%=sex%>朋友的名字</TITLE>
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
<TABLE cellSpacing=0 cellPadding=0 width=400 border=0>
<TBODY>
<TR>
<TD vAlign=top height=105>
<P>&nbsp;</P>
<P>　　　　　说明：请输入你要宴请的<%=sex%>朋友的名字。 </P></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=360 align=center border=0>
<FORM action="jiudian33.asp" method=post>

<TR>
<TD class=p9>
<div align="center">姓名：
<INPUT size=10 name=name maxlength="10">
<INPUT size=2 readonly type="hidden" name=id maxlength="2" value=<%=id%>>
</div>
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
<TD class=p9 align=middle>
<div align="center">
<INPUT type=submit value=确认>
<INPUT onclick=window.close() type=button value=关闭>
</div>
</TD></TR></FORM>
</TABLE>
<P></P>

</TD>
</TR></TBODY></TABLE></BODY></HTML>
<%
rs.close
conn.close
set rs=nothing
set conn=nothing
end if%>
