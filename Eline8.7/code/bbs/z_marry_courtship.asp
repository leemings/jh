<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_marry_conn.asp"-->
<!--#include file="z_visual_const.asp"-->
<%
stats="�������"
call nav()
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û�н��뱾����������ĵ�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
	call dvbbs_error()
else
	call marry_nav()
	call main()
end if
call activeonline()
call footer()

sub main() %>

<TABLE border=0 cellPadding=0 cellSpacing=0 width=759 align="center" height="140">
	<TR>
		<TD rowspan="2" width="26" valign="top" height="128"><img border="0" src="images/img_marry/Q01.jpg"></TD>
	    <TD colspan="4" width="740" background="images/img_marry/QB01.jpg" height="88"><img border="0" src="images/img_marry/Q02.jpg"></TD>
	</TR>
	<TR>
		<TD bgColor=#FFFFFF height="40" width="4" background="images/img_marry/QB01.jpg">��</TD>
	    <TD bgColor=#fff3ff vAlign=top height="40" width="509"> 
		<TABLE border=0 cellPadding=0 cellSpacing=0 width=487>
			<TR>
				<TD vAlign=top"></TD>
                <TD align=right vAlign=top bgcolor="#FFF3FF">
		            <TABLE align=center cellpadding=3 cellspacing=1 border=0 class=tableborder1>
						<% 
						dim rsVisual,sqlvisual,i
						set rsVisual=Server.CreateObject("adodb.recordset")
						sqlvisual="select my.username,my.fromuser,my.adddate,my.fixdate,my.period,my.price,visual.id,visual.price1,visual.price2,visual.sex,visual.content,visual.name,visual.period as period0 from visual_myitems my inner join visual_items visual on visual.id=my.itemid where my.username='"&membername&"' order by my.addDate desc"
						'set rsVisual=conn.execute("select m.*,i.name,i.sex,i.content,i.price1,i.vip from Visual_MyItems m inner join Visual_Items i on m.ItemId=i.id where m.username='"& membername &"'")
					  	rsVisual.open sqlvisual,conn,1,1
						if rsVisual.eof then 
							response.write "<tr><td class=tablebody1>���<a href=z_visual.asp?shopid=100><font color=red>�����</font></a>��������û��������<br>û��������ô�ܴ򶯹���ķ����أ���<br>�Ͽ���ȥ<a href=z_visual.asp><font color=red>�̵�</font></a>������ɣ�</td></tr>"
							'Errmsg=Errmsg+"<br>"+"<li>���<a href=z_visual.asp?shopid=100><font color=red>�����</font></a>��������û���������ȵ�<a href=z_visual.asp><font color=red>�̵�</font></a>�������"
							'call dvbbs_error()
						else 
							%><table border=0 width=100% cellPadding=0 cellSpacing=0><%
							i=0
							do while not rsVisual.eof
								i=i+1
								if (i mod 2)=1 then response.write "<tr><td>" else response.write "<td>"
								response.write "<table border=0 width=100% cellPadding=3 cellSpacing=1 class=tableborder1><tr><td rowspan=4 class=tablebody1 baseName='images/img_visual/blank.gif' width='84' height='84' border='0' itemgender='' itemno='"& rsVisual("id") &"' layerno='' localpic='"& LocalPic &"' ImagePath='"& PicPath &"' style='behavior:url(inc/z_show2.htc);'></td><td class=tablebody1><b>��Ʒ����</b>��"& rsVisual("name") &"</td></tr>"&_
												"<tr><td class=tablebody1><b>����ʱ��</b>��"& rsVisual("AddDate") &"</td></tr>"&_
												"<tr><td class=tablebody1><b>�� Ч ��</b>��"
								if rsVisual("period")=0 then
									response.write "<font color="&forum_body(8)&">��Զ��Ч</font></td>"
								elseif datediff("d",rsvisual("fixdate"),now())>rsvisual("period") then
									response.write "<font color="&forum_body(8)&">�Ѿ�����</font></td>"
								else
									response.write "��ʣ "&(rsvisual("period")-datediff("d",rsvisual("fixdate"),now()))&" ��"
								end if
									response.write "</td></tr>"&_
													"<tr><td class=tablebody1 align=center>��"
									if instr(1,"|"&v_myvisual,"|"&rsvisual("id")&".")<=0 then 
										response.write "<b><font color=blue style=""cursor:hand"" alt=""���"" onclick=""javascript:form2_"&rsvisual("id")&".myself.value=1;openScript('about:blank',500,400);form2_"&rsvisual("id")&".submit()"">���</font></b>"
									else
										response.write "<b><font color=gray style=""cursor:hand"" alt='����ʹ�õ���Ʒ������Ϊ������'>���</font></b>"
									end if
									%>��</td></tr></table>
									<form name=form2_<%=rsvisual("id")%> action=z_marry_courtshipsave.asp method=get target="openScript"><input type=hidden name=itemid value=<%=rsvisual("id")%>><input type=hidden name=myself value=0><input type=hidden name=fromuser value=""></form>
									<%
								if i=rsVisual.recordcount then
									if (i mod 2)=0 then response.write "</td></tr>" else response.write "<td colspan=2 rowspan=4>&nbsp;</td>"
								else
									if (i mod 2)=0 then	response.write "</td></tr>" else response.write "</td>" 
								end if						
								rsVisual.movenext
							loop
							response.write "</table>"
							rsVisual.close
							set rsVisual=nothing 
						end if %>
					</TABLE>
				</TD>
			</TR>
		</TABLE>
		</TD>
		<TD bgColor=#FFFFFF height="40" width="214" background="images/img_marry/QB01.jpg">��</TD>
		<TD bgColor=#FFFFFF height="40" background="images/img_marry/QB01.jpg" width="13">��</TD>
	</TR>
	<TR>
		<TD width=758 colspan="5" background="images/img_marry/QB01.jpg" height="12">��</TD>
    </TR>
</TABLE>
<% end sub %>


 