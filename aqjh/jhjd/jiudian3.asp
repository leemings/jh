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
my=aqjh_name
%><!--#include file="dadata.asp"-->
<%
sql="select * from �û� where ����='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.write "�㲻�ǽ������ˣ����ܶ�������"
conn.close
set conn=nothing
response.end
else
id=request("id")
sex=rs("�Ա�")
if sex="��" then
sex="Ů"
else
sex="��"	
end if	
%>
<HTML>
<HEAD>
<TITLE>��������Ҫ�����<%=sex%>���ѵ�����</TITLE>
<script language=JavaScript>
document.onmousedown=click;
function click(){if(event.button==2){if(confirm("�Ƿ�رմ��ڣ�","������ʾ")){window.close();}}}
</script>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK
href="regimage/home.css" type=text/css rel=stylesheet>
<STYLE>TD {
FONT-SIZE: 12px; LINE-HEIGHT: 16px
}
input {border: 1px solid; font-size: 9pt; font-family: "����"; border-color:
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
<P>����������˵������������Ҫ�����<%=sex%>���ѵ����֡� </P></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=360 align=center border=0>
<FORM action="jiudian33.asp" method=post>

<TR>
<TD class=p9>
<div align="center">������
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
<INPUT type=submit value=ȷ��>
<INPUT onclick=window.close() type=button value=�ر�>
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
