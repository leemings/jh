<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="��������"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
call main()
call activeonline()
call footer()

sub main()%>
<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
	<tr> 
		<th height=25><a name="top">����ʵʱ����</a>����<a href="#bottom">[�ײ�]</a>����<a href="z_gp_Gupiao.asp" class="cblue">[����]</a></th>
	</tr>
	<tr> 
		<td width="100%" bgcolor="#000000">
			<table border="0" width="100%">
			<tr><td colspan="4" height="10"></td></tr>
			<%set rs=server.createobject("adodb.recordset")
			sql="select content,Addtime,id from RndEvent order by ID desc"
			rs.open sql,gp_conn
			do while Not RS.Eof
				response.Write "<tr><td width=5></td><td align=left><font color=white>"&rs(2)&"</font></td><td align=left>"&rs(0)&"</td><td align=right><font color=""#FFCC99"">"&rs(1)&"</font></td><td width=5></td></tr>"
				RS.MoveNext
			Loop
			rs.close%>				
			<tr><td colspan="4" height="10"></td></tr>
			</table>
		</td> 
	</tr> 
	<tr>
		<th align=center height=25><a href="z_gp_Gupiao.asp" class="cblue">[���ع���]</a>����<a href="#top">[����]</a><a name="bottom"></a></th>
	</tr>
</table>
<%end sub%>     
