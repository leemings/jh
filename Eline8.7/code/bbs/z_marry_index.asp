<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_marry_conn.asp"-->
<%
stats="��������"
call nav()
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û�н��뱾�������õ�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
	call dvbbs_error()
else
	call marry_nav()
	call main()
end if
call activeonline()
call footer()

sub main()%>
<TABLE border=0 cellPadding=0 cellSpacing=0 width=760 align="center" height="418">
  <TBODY> 
  <TR>
    <TD vAlign=top width=519 background="images/img_marry/index_bg.jpg" height="418"><BR><BR>
      <TABLE align=center border=0 borderColor=#ff7010 cellPadding=0 
      cellSpacing=0 width=609 height="138">
        <TBODY> 
        <TR>
          <TD width="607" height="138">
            <div align="center">
              <center>
            <TABLE border=1 borderColor=#ff7010 cellPadding=0 cellSpacing=1 
            width="58%" align="right" height="75">
              <TBODY> 
              <TR>
                <TD width="15%" valign="middle" align="center" height="20"> 
                  <DIV align=center><B>׷����</B></DIV></TD>
                <TD width="15%" valign="middle" align="center" height="20"> 
                  <DIV align=center><B>������</B></DIV></TD>
                <TD width="20%" valign="middle" align="center" height="20"> 
                  <DIV align=center><B>�Ǽ�ʱ��</B></DIV></TD>	 
                <TD width="35%" valign="middle" align="center" height="20"> 
                  <DIV align=center><B>������</B></DIV></TD>			  
                <TD width="15%" valign="middle" align="center" height="20"> 
                  <DIV align=center><B>����</B></DIV></TD></TR>
				  	<% 
	set rs1=server.createobject("adodb.recordset")
    rs1.open"select * from jie ",connj,3,1
  	if rs1.bof then 
	rs1.close  %>
              <TR>
                <TD width=21% valign="middle" align="center" height="24" colspan=5>��û���κ��˵Ǽǽ������</TD>
              </tr>
                
              <% else 
		 do while not rs1.eof 
		 select case rs1("type")
		 case 1
		 str1="z_marry_cabaret.asp?name=һ��һ��&id=1" 
		 case 2
		 str1="z_marry_cabaret.asp?name=����˫��&id=2" 
		 case 3
		 str1="z_marry_cabaret.asp?name=��������&id=3" 
		 case 4
		 str1="z_marry_cabaret.asp?name=����ɽ��&id=4" 
		 case 5
		 str1="z_marry_cabaret.asp?name=����ͷ�&id=5" 
		 case 6
		 str1="z_marry_cabaret.asp?name=�������&id=6" 
		 case 7
		 str1="z_marry_cabaret.asp?name=�������&id=7" 
		 case 8
		 str1="z_marry_cabaret.asp?name=��ͷ����&id=8" 
		 case 9
		 str1="z_marry_cabaret.asp?name=�쳤�ؾ�&id=9" 
		 case else
		 str1="z_marry_cabaret.asp?name=һ��һ��&id=1" 
		 end select
		%><TR>
			  	
                <TD width="15%" valign="middle" align="center" height="19">
                  <div align="center"><a href=dispuser.asp?name=<%=rs1("username")%> target="_blank" title="�鿴<%=rs1("username")%>������"><%=rs1("username")%></a></div>
                </TD>
                <TD width="15%" valign="middle" align="center" height="19">
                  <div align="center"><a href=dispuser.asp?name=<%=rs1("thename")%> target="_blank" title="�鿴<%=rs1("thename")%>������"><%=rs1("thename")%></a></div>
                </TD>
                <TD width="20%" valign="middle" align="center" height="19">
                  <div align="center"><%=formatdatetime(rs1("addtime"),1)%></div>
                </TD>
                <TD width="35%" valign="middle" align="center" height="19">
                  <div align="left"><%=rs1("content")%></div>
                </TD>
		<%if rs1("jiehun") then%>
                <TD width="15%" align="center" valign="middle" height="19"><% if datediff("h",rs1("year"),now())>=0 and datediff("h",rs1("year"),now())<=rs1("long") then %><a href=<%=str1%> title="����Ƶ�">���ڻ���</a><%else%>�Ѿ�����<% end if %></TD>
		<%else%>
		<TD width="15%" align="center" valign="middle" height="19"><a href=z_marry_register.asp title="�Ǽǽ��">�Ǽǽ��</a></TD>
		<%end if%>
		 </TR> <% rs1.movenext
			  loop
			  rs1.close
			  set rs1=nothing 
			  end if %>
             </TBODY></TABLE></center>
            </div>
          </TD></TR></TBODY></TABLE><BR><BR></TD>
  </TR>
  </TBODY></TABLE>
	<%
end sub
%>
