<%
function iif(expression,returntrue,returnfalse)
	if expression=0 then
		iif=returnfalse
	else
		iif=returntrue
	end if
end function

'ȡ������������Ǹ�
function Max(expression1,expression2)
	if expression1+0>=expression2+0 then
		Max=expression1
	else
		Max=expression2
	end if
end function

Rem ����HTML����
function HTMLEncode(fString)
	if not isnull(fString) then
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
	
		fString = Replace(fString, CHR(32), "&nbsp;")
		fString = Replace(fString, CHR(9), "&nbsp;")
		fString = Replace(fString, CHR(34), "&quot;")
		fString = Replace(fString, CHR(39), "&#39;")
		fString = Replace(fString, CHR(13), "")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
		fString = Replace(fString, CHR(10), "<BR> ")
	
		HTMLEncode = fString
	end if
end function
'-------------------------------��Ϣ��ʾ-------------------------------
'flag=0 ��ʾ�ر�ҳ��	flag=1 ��ʾ���ع�Ʊ���״���		flag=2	��ʾ������һҳ
sub endinfo(flag) 
%>
	<tr><th valign=center align=middle height=23><b>��Ʊ������Ϣ��ʾ</b></th>
	<tr><td class=tablebody1>
<%
	if errmess<>"" then
		response.Write "<b>��������Ŀ���ԭ��</b><br><br><li>���Ƿ���ϸ�Ķ��˹���ָ�ϣ���������û�е�¼���߲�����ʹ�õ�ǰ���ܵ�Ȩ��"	
		response.Write "<font color=red>"&errmess&"</font>"
	else
		response.Write "<b>�����ɹ���</b><br><br><li>��ӭ����"&forum_info(0)&"��Ʊ�������ģ��뷵�ؽ�����������"
		response.Write "<font color=navy>"&sucmess&"</font>"
	end if		
%>	
	<br></td></tr> 
	<TR><Td class=tablebody2 height=25 align="center">
<%
	if flag=0 then
		response.Write "<a href=""../index.asp"">[������̳]</a>����<a href=""#"" onclick=""window.close()"">[�رձ�ҳ]</a>"
	else
    dim rUrl
		if rUrl="" or isnull(rurl) then
			rUrl=Request.ServerVariables("HTTP_REFERER")
		end if
		response.Write "<a href=""z_gp_GuPiao.asp"" class=cblue>[���ع��д���]</a>&nbsp;&nbsp;"	
		response.Write "<a href="""&rUrl&""" class=cblue>[������һҳ]</a>"
	end if
%></Td></TR> 
<%
end sub

sub AdminHead()%>
	<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
		<tr>
			<td align=left valign=middle height="25" class=tablebody2>&nbsp;<a href="z_gp_Gupiao.asp">��Ʊ���״���</a> �� <a href="z_gp_Admin_Setting.asp">��������</a> �� <a href="z_gp_Admin_Gupiao.asp">��Ʊ����</a> �� <a href="z_gp_Admin_User.asp">�ͻ�����</a> �� <a href="z_gp_Announcements.asp">�������</a> �� <a href="z_gp_Admin_Data.asp">���ݿ����</a> �� <a href=javascript:history.go(-1)>������һҳ</a></td>
		</tr> 
	</table>
	<br>
<%end sub

'-------------------------------��Ʊ������-------------------------------
sub CloseGuPiao(expression)
	if expression=1 then%>
		<table cellspacing=1 cellpadding=3 align=center border=0 class=tableborder1>
			<tr>
				<th height=25>��Ʊ����֪ͨ</th>
			</tr>
			<tr>
				<td width="100%" height="50" valign="middle" class=tablebody1><br>�װ��Ĺ��� <font color=blue><%=membername%></font>�����ã�<br>&nbsp;&nbsp;&nbsp;&nbsp;���ڹ�Ʊ�г��Ѿ������ˣ�����ÿ��� <%=split(Gupiao_Setting(4),"||")(0)%>:00��<%=split(Gupiao_Setting(4),"||")(1)%>:00 ������лл������<br><BR></td>
			</tr>
			<tr>
				<td align=center height=25 class=tablebody2><a href="index.asp">[������̳]</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="#" onClick="window.close()">[�رձ�ҳ]</a></td>
			</tr> 
		</table>
	<%elseif expression=0 then%>
		<table cellspacing=1 cellpadding=3 align=center border=0 class=tableborder1>
			<tr>
				<th height=25>��ӭ���� <%=Gupiao_Setting(5)%></th>
			</tr>
			<tr>
				<td width="100%" height="50" valign="middle" class=tablebody1><br><font color=<%=forum_body(8)%>>���ڹ��������ڹر�״̬�У�ԭ�����£�</font><br><br>&nbsp;&nbsp;&nbsp;<%=StopGPReadme%><br><br></td>
			</tr>
			<tr>
				<td align=center height=25 class=tablebody2><a href="index.asp">[������̳]</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="#" onClick="window.close()">[�رձ�ҳ]</a></td>
			</tr> 
		</table>
	<%else%>
		<table cellspacing=1 cellpadding=3 align=center border=0 class=tableborder1>
			<tr>
				<th height=25>��ˢ�»���</th>
			</tr>
			<tr>
				<td width="100%" height="50" valign="middle" class=tablebody1><br>&nbsp;&nbsp;��ҳ�������˷�ˢ�»��ƣ��벻Ҫ�� <font color=<%=forum_body(8)%>><%=Gupiao_Setting(7)%></font> ��������ˢ�±�ҳ��<br><br>&nbsp;&nbsp;<font color=<%=forum_body(8)%>><%=Gupiao_Setting(7)%></font> ��֮�󽫻��Զ���ҳ�棬���Ժ󡭡�<br><br></td>  
			</tr> 
			<tr>
				<td align=center height=25 class=tablebody2 id="TdReFlash"></td>
			</tr> 
		</table>
		<meta HTTP-EQUIV=REFRESH CONTENT="<%=Gupiao_Setting(7)%>">
		<script language="VBScript"> 
		<!--
			TimeLimit=<%=Gupiao_Setting(7)%>
			call GetSec()
			function GetSec()
				TimeLimit=TimeLimit-1
				TdReFlash.innerHTML = "<font color=blue>"&TimeLimit&"</font>"
				if TimeLimit<= 0 then
					TdReFlash.innerHTML="<font color=red>���ڴ�ҳ�桭��</font>"
				else
					setTimeout "GetSec()",1000 
				end if
			end function
		//-->	
		</script>		
	<%end if
end sub 

sub History()
	dim rst,TongJi,explainsplit
	set rst=gp_conn.execute("select TongJi,DangQianJiaGe,sid,ShengYuGuFen,explain from [GuPiao] order by sid")
	do while not rst.eof
		explainsplit=split(rst("explain"),"|")
		if ubound(explainsplit)<1 then
			redim preserve explainsplit(1)
			explainsplit(1)=1000
		end if
		if rst(3)>clng(explainsplit(1)) then
			sql="JiaoYiLiang="&explainsplit(1)&" "
		elseif rst(3)>0 then
			sql="JiaoYiLiang="&rst(3)
		else
			sql="JiaoYiLiang=1,ZhuangTai='��' "
		end if	
		TongJi=rst(0)
		TongJi=mid(TongJi,instr(TongJi,"|")+1)&"|"&formatnumber(rst(1),2,-1)
		gp_conn.execute("update [GuPiao] set TongJi='"&TongJi&"',KaiPanJiaGe=DangQianJiaGe,ChengJiaoLiang=0,MaiRuBiShu=0,MaiChuBiShu=0,TodayWave=0,"&sql&" where sid="&rst(2))
		rst.movenext
	loop
end sub
%>