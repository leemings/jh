<!--#include file="conn.asp"-->
<!--#include file="Z_TestConn.ASP"-->
<!-- #include file="inc/const.asp" -->
<meta http-equiv="Content-Language" content="zh-cn">
<%
'-----------------------
'edit by ������/��ˮ��ɽ
'-----------------------
	dim ars

	stats="���Ĵǵ� �ǵ����"
			
	if not master then
		errmsg=errmsg+"<br><li>��û�й����Ĵǵ��Ȩ�ޣ���ȷ�����Ѿ���¼���Ҿ�����Ӧ��Ȩ�ޣ�"
		stats="���Ĵǵ� ������Ϣ"
		call nav()
		call head_var(0,0,"���Ĵǵ�","Z_test.asp")		
		call dvbbs_error()	
	else
		call nav()
		call head_var(0,0,"���Ĵǵ�","Z_test.asp")
%>		
		<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
		<tr><th>��ӭ <%=membername%> ���뿪�Ĵǵ� ��������</th></tr>
		<tr><td class=tablebody2>
<%	
		call Admin_Head()
			
		select case Request("action") 
			case "setting"
				call setting()
			case "OtherSetting"
				call OtherSetting()	
			case "editlb"
				call editlb()
			case "dellb"
				call dellb()
			case "addlb"
				call addlb()
			case "move"
				call move()	
			case else
				call main()
		end select
	end if
%>
		<br>
	</td></tr>
	</table>
<%	
	set ars=nothing
	set aconn=nothing	
	call activeonline()
	call footer()
	
'=================================================						
sub main()			 
	set ars=aconn.execute("select * from [config]")
	dim AnswerSetting,OtherSetting
	AnswerSetting=split(ars("AnswerSetting"),"||")
	if ubound(AnswerSetting)<8 then
		redim preserve AnswerSetting(8)
		AnswerSetting(8)="1"
	end if
	OtherSetting=split(ars("OtherSetting"),"||")	
%>
	<table align=center cellspacing=1 cellpadding="3" class=tableborder1 style="width:97%">
		<tr>
	    	<td align="left" class=TopLighNav1> �� <b><a name=top>��������</a></b> <a href="#middle">[�м�]</a> <a href="#bottom">[�׶�]</a></td>
		  	<td align="left" class=TopLighNav1> �� <b>��Ŀ������</b> <a href="#middle">[�м�]</a> <a href="#bottom">[�׶�]</a></td>
		</tr>
		<tr>
		<td width="50%" valign="top" class=tablebody1>
		<form method="POST" action="Z_TestSetting.asp?action=setting" name=setting>
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
			
				<tr><td align="left" class=tablebody2 colspan="2">�� ÿ��һ��������</td></tr>
				<tr>
					<td width="40%" align="right" class=tablebody1>��Ϸ�ң�</td>
					<td width="60%" align="left" class=tablebody1><font color=<%=Forum_body(8)%>> �� </font><input type="text" name="userp" size="10" value=<%=AnswerSetting(0)%>> ��</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>����ֵ��</td>
					<td align="left" class=tablebody1><font color=<%=Forum_body(8)%>> �� </font><input type="text" name="userc" size="10" value=<%=AnswerSetting(1)%>> ��</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>����ֵ��</td>
					<td align="left" class=tablebody1><font color=<%=Forum_body(8)%>> �� </font><input type="text" name="usere" size="10" value=<%=AnswerSetting(2)%>> ��</td>
				</tr>
				
				<tr><td align="left" class=tablebody2 colspan="2">�� �������ã�</td></tr>
				<tr>
					<td align="right" class=tablebody1>ÿ�δ�����ʱ�䣺</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="dttime" size="10" value=<%=AnswerSetting(3)%>> ����</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>ÿ������������</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="ts" size="10" value=<%=AnswerSetting(4)%>> ��</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>��ʾ��ȷ�𰸣�</td> 
					<td align="left" class=tablebody1>&nbsp; <input type="radio" name="ShowAnswer" value=1 <%if cint(AnswerSetting(5))=1 then%>checked<%end if%>> ��&nbsp;&nbsp;<input type="radio" name="ShowAnswer" value=0 <%if cint(AnswerSetting(5))=0 then%>checked<%end if%>> ��</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>��Ǯ��֪ʶ�����Ǯ��</td> 
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="AnswerMoney" size="10" value=<%=AnswerSetting(6)%>> ��</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>�����Ƿ���ʱ��</td> 
					<td align="left" class=tablebody1>&nbsp; <input type="radio" name="TimeLimit" value=1 <%if cint(AnswerSetting(7))=1 then%>checked<%end if%>> ��&nbsp;&nbsp;<input type="radio" name="TimeLimit" value=0 <%if cint(AnswerSetting(7))=0 then%>checked<%end if%>> ��</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>��Ŀ�Ƿ�֧��UBB��</td>
					<td align="left" class=tablebody1>&nbsp; <input type="radio" name="UBB" value=1 <%if cint(AnswerSetting(8))=1 then%>checked<%end if%>> ��&nbsp;&nbsp;<input type="radio" name="UBB" value=0 <%if cint(AnswerSetting(8))=0 then%>checked<%end if%>> ��</td>
				</tr>
				
				<tr><td align="left" class=tablebody2 colspan="2">�� <a name=middle>�������ã�</a></td></tr>				
				<tr>
					<td align="right" class=tablebody1>�û��ϴ���Ŀ������</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="userjl" size="10" value=<%=ars("userjl")%>> Ԫ</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>�ϴ�һ�⽱��Ϸ�ң�</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="userupsinew" size="10" value=<%=ars("userupsinew")%>> ö</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>ÿ��һ������Ϸ�ң�</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="userarticle" size="10" value=<%=ars("userarticle")%>> ö</td>
				</tr>								
				<tr>
					<td align="right" class=tablebody1>��ɴ��⸽�ӽ�����</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="fjjl" size="10" value=<%=ars("fjjl")%>> Ԫ</td>
				</tr>
				
				<tr><td align="left" class=tablebody2 colspan="2">�� ��ø��ӽ���������</td></tr>				
				<tr>
					<td align="right" class=tablebody1>�������ٴ�������</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="dts" size="10" value=<%=ars("dts")%>> ��</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>���δ�����ȷ�ʣ�</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="zql" size="10" value=<%=ars("zql")%>> %</td>
				</tr>
				
				<tr><td align="left" class=tablebody2 colspan="2">�� �ϴ����ã�</td></tr>				
				<tr>
					<td align="right" class=tablebody1>�����ϴ�ѡ���⣺</td> 
					<td align="left" class=tablebody1>&nbsp; <input type="radio" name="CanUpChoose" value=1 <%if cint(ars("CanUpChoose"))=1 then%>checked<%end if%>> ��&nbsp;&nbsp;<input type="radio" name="CanUpChoose" value=0 <%if cint(ars("CanUpChoose"))=0 then%>checked<%end if%>> ��</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>�����ϴ�����⣺</td>
					<td align="left" class=tablebody1>&nbsp; <input type="radio" name="CanUpFill" value=1 <%if cint(ars("CanUpFill"))=1 then%>checked<%end if%>> ��&nbsp;&nbsp;<input type="radio" name="CanUpFill" value=0 <%if cint(ars("CanUpFill"))=0 then%>checked<%end if%>> ��</td>
				</tr>								
				<tr>
					<td align="right" class=tablebody1>��Ŀ���������</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="yhuptitlebyte" size="10" value=<%=ars("yhuptitlebyte")%>> Byte</td>
				</tr>
				<tr>
					<td align="right" class=tablebody1>�����������</td>
					<td align="left" class=tablebody1>&nbsp;&nbsp; <input type="text" name="yhupanswbyte" size="10" value=<%=ars("yhupanswbyte")%>> Byte</td>
				</tr>								
			</table>
			
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<tr><td align="center" class=tablebody2><input type="submit" value="�ύ" name="submit">&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" value="����" name="reset"></td></tr>
			</table>
			
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<tr><td align="left" class=tablebody2>
				������ʾ��
				<br>&nbsp;�� <font color="#FF0000">ע�⣺�������õĲ�������Ϊ���֣����Ҳ���Ϊ����Ŷ</font>
				<br>&nbsp;�� ����Ϊ��ǰִ�еĲ���
				<br>&nbsp;�� ���"����"�ָ���ǰִ�еĲ���
				<br>&nbsp;�� ������ʱ����Ϊ��ʱ����Ϊ���ⲻ���Ƽ��ʱ��				
				</td></tr>
			</table>
			</form>					
		</td>
		
		<td valign="top" class=tablebody2 width="50%">
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<tr>
					<td align="center" class=TopLighNav1 width="40">�򣠺�</td>
					<td align="center" class=TopLighNav1 width=*>�࣠����</td>
					<td align="center" class=TopLighNav1 width="100">�٣���</td>
				</tr>		
<% 
				sql="select lb,id from [testlb]"
				set rs=aconn.execute(sql)
				if rs.eof then
					response.write "<tr><td class=tablebody1 colspan=3>��ʱ��û��������</td></tr>"
				else
					do while not rs.eof
						response.write "<tr><td align=center class=tablebody1><font color='#FF0000'>"&rs(1)&"</font></td><td align=center class=tablebody1>"&rs(0)&"</td>"
						response.write "<td align=center class=tablebody1><a href='Z_TestSetting.asp?action=dellb&id="&rs(1)&"'"&" onclick=""javascript:{if(confirm('ɾ������������ͬ����ڴ�������ͬʱɾ������ȷ��ִ��ɾ��������?')){return true;}return false;}""><font color='#FF0000'>ɾ��</a></font></td></tr>"
						rs.movenext
					loop
					rs.close
					set rs=nothing
				end if	
%>
			</table>
			
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<form method="POST" action="Z_TestSetting.asp?action=editlb" name=form3>
				<tr><td align="left" class=TopLighNav1>�� �޸������</td></tr>
				<tr><td align="left" class=tablebody1>&nbsp;��ţ�<input type="text" name="id" size="3">&nbsp;&nbsp;&nbsp;������ƣ�<input type="text" name="LbName" size="10">&nbsp;&nbsp;&nbsp;&nbsp;	<input type="submit" value="�޸�" name="B3"></td></tr>
				</form>
			</table>
			
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<form method="POST" action="Z_TestSetting.asp?action=addlb" name=form4>
				<tr><td align="left" class=TopLighNav1>�� �������</td></tr> 
				<tr><td align="left" class=tablebody1>&nbsp;������ƣ�<input type="text" name="NewLbName" size="10">&nbsp;&nbsp;&nbsp;<input type="submit" value="���" name="B4"></td></tr>
				</form>
			</table>

			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<form method="POST" action="Z_TestSetting.asp?action=move" name=form5>
				<tr><td align="left" class=TopLighNav1>�� ��Ŀ����ת��</td></tr>
				<tr><td align="left" class=tablebody1>&nbsp;Դ�����ţ�<input type="text" name="OutID" size="3">&nbsp;&nbsp;&nbsp;Ŀ�������ţ�<input type="text" name="InID" size="3">&nbsp;&nbsp;&nbsp;&nbsp;	<input type="submit" value="ת��" name="B4"></td></tr>
				</form>
			</table>
									  
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<tr><td align="left" class=tablebody2>
				������ʾ��
				<br>&nbsp;�� <font color="#FF0000">ɾ�����</font>��������ͬ����ڴ�������ͬʱɾ����������
				<br>&nbsp;�� <font color="#FF0000">�޸����</font>���ƿ�ֱ�����������������������Ӧ��������ƺ���<font color="#FF0000">�޸�</font>����  
				<br>&nbsp;�� <font color="#FF0000">�������</font>������Ŀ�������������Ƶ��<font color="#FF0000">���</font>����
				<br>&nbsp;�� <font color="#FF0000">����ת��</font>�������ǰ�����ĳ�������������Ŀ�Ƶ�����һ��������
				</td></tr>
			</table>
			
			</td>
		</tr>

		<tr>
	    <td align="left" class=TopLighNav1> �� <b><a name=bottom>��������</a></b> <a href="#middle">[�м�]</a> <a href="#top">[����]</a></td>
		<td align="center" class=TopLighNav1><b></b></td>
		</tr>
		<form method="POST" action="Z_TestSetting.asp?action=OtherSetting" name=form2>
		<tr>
		<td width="50%" valign="top" class=tablebody1>
			
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
											
				<tr>
					<td align="left" class=tablebody1 width="45%">�ǵ���ҳ�Ƿ���ʾͼƬ��</td> 
					<td align="left" class=tablebody1 width="55%">&nbsp;<input type="radio" name="ShowPic" value=1 <%if cint(OtherSetting(0))=1 then%>checked<%end if%>> ͼƬ&nbsp;&nbsp;<input type="radio" name="ShowPic" value=2 <%if cint(OtherSetting(0))=2 then%>checked<%end if%>> ���&nbsp;&nbsp;<input type="radio" name="ShowPic" value=3 <%if cint(OtherSetting(0))=3 then%>checked<%end if%>> ������Ϣ</td>
				</tr>
				<tr>
					<td align="left" class=tablebody1>�ǵ���ҳͼƬ·����</td> 
					<td align="left" class=tablebody1>&nbsp; <input type="text" name="PicPath" value="<%=OtherSetting(1)%>" size="25"></td>
				</tr>				
				<tr>
					<td align="left" class=tablebody1>������������ʾ��¼������</td> 
					<td align="left" class=tablebody1>&nbsp; <input type="text" name="PaiXingTop" value="<%=OtherSetting(2)%>"> ��</td>  
				</tr>				
			</table>
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<tr><td align="center" class=tablebody2><input type="submit" value="�ύ" name="submit">&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" value="����" name="reset"></td></tr>
			</table>			
			
		</td>
		<td width="50%" valign="top" class=tablebody1>
			<table border=0><tr><td height="3"></td></tr></table>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
				<tr><td align="left" class=TopLighNav1>�� �ǵ���ҳ������[<font color=red>֧��HTML</font>]</td></tr>							
				<tr>
					<td align="left" class=tablebody1 width="100%">
					<textarea name="kxcd_ad" cols="65" rows="5"><%=ars("ad")%></textarea> 
					</td>
				</tr>
				
			</table>
			<table border=0><tr><td height="3"></td></tr></table>		
		</td>
		</tr>
		</form>		
		</table>
<%	
	ars.close
end sub

sub setting()
	if request("userp")="" or (not isnumeric(request("userp"))) then
		errmsg=errmsg+"<br><li>��Ϸ�Ҳ���Ϊ�ջ��߷����֣�����������"
		founderr=true
	elseif request("userp")<0 then 
		errmsg=errmsg+"<br><li>��Ϸ�Ҳ���Ϊ����������������"
		founderr=true	
	end if
	if request("userc")="" or (not isnumeric(request("userc"))) then
		errmsg=errmsg+"<br><li>����ֵ����Ϊ�ջ��߷����֣�����������"
		founderr=true
	elseif request("userc")<0 then 
		errmsg=errmsg+"<br><li>����ֵ����Ϊ����������������"
		founderr=true	
	end if
	if request("usere")="" or (not isnumeric(request("usere"))) then
		errmsg=errmsg+"<br><li>����ֵ����Ϊ�ջ��߷����֣�����������"
		founderr=true
	elseif request("usere")<0 then 
		errmsg=errmsg+"<br><li>����ֵ����Ϊ����������������"
		founderr=true	
	end if	
	if request("dttime")="" or (not isnumeric(request("dttime"))) then
		errmsg=errmsg+"<br><li>������ʱ�䲻��Ϊ�ջ��߷����֣�����������"
		founderr=true
	elseif request("dttime")<0 then 
		errmsg=errmsg+"<br><li>������ʱ�䲻��Ϊ����������������"
		founderr=true	
	end if
	if request("ts")="" or (not isnumeric(request("ts"))) then
		errmsg=errmsg+"<br><li>ÿ��������������Ϊ�ջ��߷����֣�����������"
		founderr=true
	elseif request("ts")<0 then 
		errmsg=errmsg+"<br><li>ÿ��������������Ϊ����������������"
		founderr=true	
	end if
	if request("AnswerMoney")="" or (not isnumeric(request("AnswerMoney"))) then
		errmsg=errmsg+"<br><li>��Ǯ��֪ʶ�����Ǯ����Ϊ�ջ��߷����֣�����������"
		founderr=true
	elseif request("AnswerMoney")<0 then 
		errmsg=errmsg+"<br><li>��Ǯ��֪ʶ�����Ǯ����Ϊ����������������"
		founderr=true	
	end if
	if request("userjl")="" or (not isnumeric(request("userjl"))) then
		errmsg=errmsg+"<br><li>�û��ϴ���Ŀ��������Ϊ�ջ��߷����֣�����������"
		founderr=true
	elseif request("userjl")<0 then 
		errmsg=errmsg+"<br><li>�û��ϴ���Ŀ��������Ϊ����������������"
		founderr=true	
	end if
	if request("fjjl")="" or (not isnumeric(request("fjjl"))) then
		errmsg=errmsg+"<br><li>��ɴ��⸽�ӽ�������Ϊ�ջ��߷����֣�����������"
		founderr=true
	elseif request("fjjl")<0 then 
		errmsg=errmsg+"<br><li>��ɴ��⸽�ӽ�������Ϊ����������������"
		founderr=true	
	end if
	if request("zql")="" or (not isnumeric(request("fjjl"))) then
		errmsg=errmsg+"<br><li>���δ�����ȷ�ʲ���Ϊ�ջ��߷����֣�����������"
		founderr=true
	elseif request("zql")<0 or request("zql")>100 then 
		errmsg=errmsg+"<br><li>���δ�����ȷ�ʲ��ܴ���100����С��0������������"
		founderr=true	
	end if
	if request("dts")="" or (not isnumeric(request("dts"))) then
		errmsg=errmsg+"<br><li>�������ٴ���������Ϊ�ջ��߷����֣�����������"
		founderr=true
	elseif request("dts")<0 then 
		errmsg=errmsg+"<br><li>�������ٴ���������Ϊ����������������"
		founderr=true	
	end if
	if request("userupsinew")="" or (not isnumeric(request("userupsinew"))) then
		errmsg=errmsg+"<br><li>ÿ��һ������Ϸ�Ҳ���Ϊ�ջ��߷����֣�����������"
		founderr=true
	elseif request("userupsinew")<0 then 
		errmsg=errmsg+"<br><li>ÿ��һ������Ϸ�Ҳ���Ϊ����������������"
		founderr=true	
	end if
	if request("userarticle")="" or (not isnumeric(request("userarticle"))) then
		errmsg=errmsg+"<br><li>�ϴ�һ�⽱��Ϸ�Ҳ���Ϊ�ջ��߷����֣�����������"
		founderr=true
	elseif request("userarticle")<0 then 
		errmsg=errmsg+"<br><li>�ϴ�һ�⽱��Ϸ�Ҳ���Ϊ����������������"
		founderr=true	
	end if
	if request("yhuptitlebyte")="" or (not isnumeric(request("yhuptitlebyte"))) then
		errmsg=errmsg+"<br><li>��Ŀ�����������Ϊ�ջ��߷����֣�����������"
		founderr=true
	elseif request("yhuptitlebyte")<0 then 
		errmsg=errmsg+"<br><li>��Ŀ�����������Ϊ����������������"
		founderr=true	
	end if
	if request("yhupanswbyte")="" or (not isnumeric(request("yhupanswbyte"))) then
		errmsg=errmsg+"<br><li>�������������Ϊ�ջ��߷����֣�����������"
		founderr=true
	elseif request("yhupanswbyte")<0 or request("yhupanswbyte")>200then 
		errmsg=errmsg+"<br><li>�������������С��0���ߴ���200������������"
		founderr=true	
	end if										
	
	if founderr then
		call test_err()
		exit sub
	end if
	
	if cint(request("dts"))>cint(request("ts")) then
		errmsg=errmsg+"<br><li>�ۣ����ˣ��������ٴ����� �� ������������ ��Ҫ�డ��˭���õ�����Ŷ��"
		call test_err()
		exit sub
	end if		
	
	dim AnswerSetting 
	AnswerSetting=request("userp")&"||"&request("userc")&"||"&request("usere")&"||"&request("dttime")&"||"&request("ts")&"||"&request("ShowAnswer")&"||"&request("AnswerMoney")&"||"&request("TimeLimit")&"||"&request("UBB")
	aconn.execute("update [config] set AnswerSetting='"&AnswerSetting&"',userjl="&request("userjl")&",fjjl="&request("fjjl")&",zql="&request("zql")&",dts="&request("dts")&",userupsinew="&request("userupsinew")&",userarticle="&request("userarticle")&",yhuptitlebyte="&request("yhuptitlebyte")&",yhupanswbyte="&request("yhupanswbyte")&",CanUpChoose="&request("CanUpChoose")&",CanUpFill="&request("CanUpFill")&"")
	sucmsg=sucmsg+"<br><li>�������óɹ����뷵�ؽ�����������"
	call suc()
end sub

sub OtherSetting()
	if request("ShowPic")=1 and request("PicPath")="" then
		errmsg=errmsg+"<br><li>������ͼƬ·��"
		founderr=true
	end if
	if request("PaiXingTop")="" then
		errmsg=errmsg+"<br><li>�����������������ʾ��¼����"
		founderr=true	
	elseif not isnumeric(request("PaiXingTop")) then
		errmsg=errmsg+"<br><li>������������ʾ��¼����������������"
		founderr=true
	elseif request("PaiXingTop")<=0 then
		errmsg=errmsg+"<br><li>������������ʾ��¼�������������"
		founderr=true
	end if	
	if founderr then
		call test_err()
		exit sub
	end if	
			
	dim OtherSetting  
	OtherSetting=request("ShowPic")&"||"&request("PicPath")&"||"&request("PaiXingTop")
	aconn.execute("update [config] set OtherSetting='"&OtherSetting&"',ad='"&CheckStr(request("kxcd_ad"))&"'")
	sucmsg=sucmsg+"<br><li>�������óɹ����뷵�ؽ�����������"
	call suc()
end sub

sub editlb()
	if request("id")="" or isnull(request("id")) then
		errmsg=errmsg+"<br><li>������Ҫ�޸ĵ��������"
		founderr=true
	elseif not isnumeric(request("id")) then
		errmsg=errmsg+"<br><li>��ű���������"	
		founderr=true
	end if	
	if request("LbName")="" then
		errmsg=errmsg+"<br><li>���������������"
		founderr=true
	end if 
	if founderr then
		call test_err()
		exit sub
	end if
    set ars=aconn.execute("select lb from [testlb] where id="&request("id"))
	if ars.eof and ars.bof then
		ars.close
		errmsg=errmsg+"<br><li>û����������������ȷ��������"
		founderr=true
		call test_err()
		exit sub		
	elseif trim(ars(0))=trim(request("LbName")) then
		ars.close
		errmsg=errmsg+"<br><li>���������������Ѿ����ڣ�����������"
		founderr=true
		call test_err()
		exit sub						
	end if
	ars.close
			
	aconn.execute("update [testlb] set lb='"&trim(replace(request("LbName"),"'",""))&"' where id="&request("id"))
	
	sucmsg=sucmsg+"<br><li>������޸ĳɹ����뷵�ؽ�����������"
	call suc()	
end sub

sub dellb()
	aconn.execute ("delete from [testlb] where id="&request("id"))
	aconn.execute ("delete from [test] where lb="&request("id"))
	sucmsg=sucmsg+"<br><li>���ɾ���ɹ������ڸ�������ĿҲ�Ѿ�ɾ�����뷵�ؽ�����������"
	call suc()
end sub

sub addlb()
	if request("NewLbName")="" then
		errmsg=errmsg+"<br><li>�������������"
		call test_err()
		exit sub
	end if 
	
	set ars=aconn.execute("select lb from [testlb] where lb='"&trim(replace(request("NewLbName"),"'",""))&"'")
	if not (ars.eof and ars.bof) then
		errmsg=errmsg+"<br><li>���������������Ѿ����ڣ�����������"
		founderr=true
		ars.close
		call test_err()
		exit sub						
	end if
	ars.close		
	aconn.execute("insert into [testlb](lb) values('"&trim(replace(request("NewLbName"),"'",""))&"')")
	
	sucmsg=sucmsg+"<br><li>�ɹ�����������뷵�ؽ�����������"
	call suc()
end sub

sub move()
	if request("OutID")="" then
		errmsg=errmsg+"<br><li>�������Ƴ��ķ�������"
		founderr=true
	elseif not isInteger(request("OutID")) then
		errmsg=errmsg+"<br><li>�Ƴ��ķ������Ų���ȷ������������"
		founderr=true		
	end if 
	if request("InID")="" then
		errmsg=errmsg+"<br><li>����������ķ�������"
		founderr=true
	elseif not isInteger(request("InID")) then
		errmsg=errmsg+"<br><li>����ķ������Ų���ȷ������������"
		founderr=true		
	end if 
		
	if founderr then
		call test_err()
		exit sub
	end if
		
	set ars=aconn.execute("select id from [testlb] where id="&request("InID"))
	if ars.eof and ars.bof then
		errmsg=errmsg+"<br><li>Ŀ����𲻴���"
		founderr=true
	end if
	ars.close
	
	set ars=aconn.execute("select id from [test] where lb="&request("OutID"))
	if ars.eof and ars.bof then
		errmsg=errmsg+"<br><li>��Դ�����û���κ���Ŀ����Դ��𲢲�����"
		founderr=true
	end if
	ars.close	
	
	if founderr then 
		call test_err()
	else	
		aconn.execute("update [test] set lb="&request("InID")&" where lb="&request("OutID"))
		sucmsg=sucmsg+"<br><li>�����ɹ����뷵�ؽ�����������"
		call suc()
	end if
end sub
%>