<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head><title><%=Gupiao_Setting(5)%>-���п������</title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<table  width="98%" border=0 cellspacing=1 cellpadding=0 align=center bgcolor="#0066CC">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
	</TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif"><font size="2"><b>���й�������</b></font></td>
	</tr>
	<tr><td bgcolor="#E4E8EF"> 
<% 
call AdminHead()
if not master then
	errmess="<li>��ҳΪ����Աר�ã���û�й���ҳ��Ȩ�ޣ�"
	call endinfo(1)
else 
	select case request("action")
		case "SaveSettimg"		
			call SaveSettimg()
		case else
			call main()
	end select
end if
%>
<br>
</td></tr>
<tr><td height=32 background="images/footer.gif" valign=middle></td></tr>
</table>
</body></html>
<%
CloseDatabase		'�ر����ݿ�
'=====================================
sub main()
	dim KaiHu_Setting,PStock_Setting,User_Setting
	set rs=conn.execute("select top 1 PStock_Setting,KaiHu_Setting,Trade_Setting,User_Setting,AI_Setting,Custom_Setting from [GupiaoConfig] order by id") 
	PStock_Setting=split(rs("PStock_Setting"),"|")
	KaiHu_Setting =split(rs("KaiHu_Setting"),"|")
	'Trade_Setting =split(rs("Trade_Setting"),"|")
	User_Setting  =split(rs("User_Setting"),"|")
	'Custom_Setting=split(rs("Custom_Setting"),"||")
	
%>
<table width="97%" border="0" cellspacing="1" cellpadding="3" align="center" style="border: 1px #6595D6 solid; background-color: #FFFFFF;">
	<tr>
		<td align="center" colspan="2" height="25" background="Images/Header.gif"><a name=top></a><b>��Ʊ����</b></td>
	</tr>
	<tr>
		<td class="forumRow" colspan="2">
<table width=85% align=center>
<tr>
<td width=50% ><a href="#Gupiao_Setting">[��������]</a></td>
<td width=50% ><a href="#Gupiao_Setting1">[��ˢ�»���]</a></td>
</tr>
<tr>
<td width=50% ><a href="#Gupiao_Setting2">[����¼�]</a></td>
<td width=50% ><a href="#Trade_Setting">[��������]</a></td>
</tr>
<tr>
<td width=50% ><a href="#PStock_Setting">[������������]</a></td>
<td width=50% ><a href="#AI_Setting">[������(AI)����]</a></td>
</tr>
<tr>
<td width=50% ><a href="#KaiHu_Setting">[��������]</a></td>
<td width=50% ><a href="#User_Setting">[�ʻ�����]</a></td>
</tr>
<tr>
<td width=50% ><a href="#Custom_Setting">[�Զ���λ��]</a></td>
<td width=50% ><a href="#Rule_Setting">[���й�������]</a></td>
</tr>
</table>		
		</td>
	</tr>	
		
	<form method="POST" action="?action=SaveSettimg">
	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=Gupiao_Setting></a>�������� [<a href="#top">����</a>] [<a href="#bottom">�ײ�</a>]</th>
	</tr>	
	<tr>
		<td class="forumRow" width="40%"><b>��Ʊ�г�״̬</b><br>�����Ƿ�ͨ</td> 
		<td class="forumRow" width="60%">
			<input type=radio name="Gupiao_Setting(0)" value=1 <%if cint(Gupiao_Setting(0))=1 then%>checked<%end if%>>��&nbsp;
			<input type=radio name="Gupiao_Setting(0)" value=0 <%if cint(Gupiao_Setting(0))=0 then%>checked<%end if%>>�ر�&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow" width="40%"><b>�ر�˵��</b><br>�ڹ��йرյ��������ʾ��֧��html�﷨</td> 
		<td class="forumRow" width="60%">
			<textarea name="StopReadme" cols="40" rows="3"><%=StopReadme%></textarea>
		</td>
	</tr>	
	<tr>
		<td class="forumRow"><b>�Զ�ˢ��ʱ��(��)</b><br>������ҳ�Զ�ˢ�µ�ʱ����</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Gupiao_Setting(1)" value="<%=Gupiao_Setting(1)%>"> ��
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>ÿҳ��ʾ����¼</b><br>���ڹ�Ʊ���кͷ�ҳ�йص���Ŀ</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Gupiao_Setting(2)" value="<%=Gupiao_Setting(2)%>"> ��
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>�Ƿ����ö�ʱ���ع���</b></td> 
		<td class="forumRow">
			<input type=radio name="Gupiao_Setting(3)" value=1 <%if cint(Gupiao_Setting(3))=1 then%>checked<%end if%>>��&nbsp;
			<input type=radio name="Gupiao_Setting(3)" value=0 <%if cint(Gupiao_Setting(3))=0 then%>checked<%end if%>>��&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>���п���ʱ��</b><br>��ȷ�����Ѿ��������ö�ʱ���ع���<br>��ֹСʱ�÷��š�||���ֿ�</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Gupiao_Setting(4)" value="<%=Gupiao_Setting(4)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>��������</b></td> 
		<td class="forumRow">
			&nbsp;<input type=text size="35" name="Gupiao_Setting(5)" value="<%=Gupiao_Setting(5)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><a name=Gupiao_Setting1></a><b>��ˢ�»���</b><br>��ѡ�������д���������ˢ��ʱ��</td> 
		<td class="forumRow">
			<input type=radio name="Gupiao_Setting(6)" value=1 <%if cint(Gupiao_Setting(6))=1 then%>checked<%end if%>>��&nbsp;
			<input type=radio name="Gupiao_Setting(6)" value=0 <%if cint(Gupiao_Setting(6))=0 then%>checked<%end if%>>�ر�&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>���ˢ��ʱ����</b><br>��д����Ŀ��ȷ�������˷�ˢ�»���</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Gupiao_Setting(7)" value="<%=Gupiao_Setting(7)%>"> ��
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>��ˢ�¹�����Ч��ҳ��</b><br>��ȷ�������˷�ˢ�¹���<br>��ָ����ҳ�潫�з�ˢ�����ã��û����޶���ʱ���ڲ����ظ��򿪸�ҳ�棬����һ��������Դ���ĵ�����<br>ÿ��ҳ�������á�|�����Ÿ���</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="35" name="Gupiao_Setting(8)" value="<%=Gupiao_Setting(8)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>ɾ������û�ʱ��</b><br>������ɾ�����ٷ����ڲ���û�<br>��λ�����ӣ�����������</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="5" name="Gupiao_Setting(9)" value="<%=Gupiao_Setting(9)%>"> ����
		</td>
	</tr>
	<tr>
		<td class="forumRow"><a name=Gupiao_Setting2></a><b>����¼���������</b><br>������(0��1)֮����ֵ</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="5" name="Gupiao_Setting(10)" value="<%=Gupiao_Setting(10)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>����¼��ɼ����ǵļ���</b><br>��������¼�ʱ�ɼ����ǵļ���<br>�ɼ��½��ļ���=1-�ɼ����ǵļ���</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="5" name="Gupiao_Setting(11)" value="<%=Gupiao_Setting(11)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>��������¼���¼������</b><br>����Ϊ0������������¼���¼</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="Gupiao_Setting(12)" value="<%=Gupiao_Setting(12)%>"> ��
		</td>
	</tr>
		
	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=Trade_Setting></a>��Ʊ�������� [<a href="#top">����</a>] [<a href="#bottom">�ײ�</a>]</th>
	</tr>
	<tr>
		<td class="forumRow"><b>�Ƿ�������</b><br>ֻ�б����ʱ������ܽ��й�Ʊ��������</td> 
		<td class="forumRow">
			<input type=radio name="Trade_Setting(0)" value=1 <%if cint(Trade_Setting(0))=1 then%>checked<%end if%>>��&nbsp;
			<input type=radio name="Trade_Setting(0)" value=0 <%if cint(Trade_Setting(0))=0 then%>checked<%end if%>>��&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>�ɽ�������</b><br>�����������ٽ�����<br>��������������Ϊ0</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(2)" value="<%=Trade_Setting(2)%>"> ��		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>�ɽ�������</b><br>�����������ཻ����<br>��������������Ϊ0</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(3)" value="<%=Trade_Setting(3)%>"> ��		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>���״�������</b><br>ͬһ�û�ÿ�콻�״�������<br>��������������Ϊ0</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(4)" value="<%=Trade_Setting(4)%>"> ��		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>������</b><br>ÿ�ʽ�����ȡ�������ѵİٷ���</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(5)" value="<%=Trade_Setting(5)%>">  %	
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>����ʱ������</b><br>ͬһ�û�����ͬһ�ֹ�Ʊ��ʱ����</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(6)" value="<%=Trade_Setting(6)%>"> ����
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>����ʱ������</b><br>ͬһ�û�����ͬһ�ֹ�Ʊ��ʱ����</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(7)" value="<%=Trade_Setting(7)%>"> ����
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>����ʱ������</b><br>ͬһ�û�����ĳ��Ʊ����������ʱ����</td> 
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(8)" value="<%=Trade_Setting(8)%>"> Сʱ
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>��ͣ����</b><br>��Ʊ�۸������ʴ��ڸ�ֵʱ������ͣ״̬</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(9)" value="<%=Trade_Setting(9)%>">
		</td>
	</tr>	
	<tr>
		<td class="forumRow"><b>��ͣ����</b><br>��Ʊ�۸��»��ʴ��ڸ�ֵʱ���ֵ�ͣ״̬</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(10)" value="<%=Trade_Setting(10)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>ͣ������</b><br>��Ʊ�۸���ڸ�ֵʱ����ͣ��״̬</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="Trade_Setting(11)" value="<%=Trade_Setting(11)%>"> Ԫ
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>����Թɼ۵�Ӱ��</b><br>�����Ʊʱ����Ʊ�۸� ���ǡ��½�������ı���<br>����+�½�<=1,����ļ���=1-����-�½�</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="5" name="Trade_Setting(12)" value="<%=Trade_Setting(12)%>">��<input type=text size="5" name="Trade_Setting(13)" value="<%=Trade_Setting(13)%>">��<%=formatnumber(1-Trade_Setting(12)-Trade_Setting(13),2,true)%>
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>�����Թɼ۵�Ӱ��</b><br>�׳���Ʊʱ����Ʊ�۸� �½������ǡ�����ı���<br>����+�½�<=1,����ļ���=1-����-�½�</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="5" name="Trade_Setting(14)" value="<%=Trade_Setting(14)%>">��<input type=text size="5" name="Trade_Setting(15)" value="<%=Trade_Setting(15)%>">��<%=formatnumber(1-Trade_Setting(14)-Trade_Setting(15),2,true)%>
		</td>
	</tr>
																			
	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=PStock_Setting></a>������������ [<a href="#top">����</a>] [<a href="#bottom">�ײ�</a>]</th>
	</tr>
	<tr>
		<td class="forumRow"><b>�Ƿ���������������</b></td> 
		<td class="forumRow">
			<input type=radio name="PStock_Setting(0)" value=1 <%if cint(PStock_Setting(0))=1 then%>checked<%end if%>>��&nbsp;
			<input type=radio name="PStock_Setting(0)" value=0 <%if cint(PStock_Setting(0))=0 then%>checked<%end if%>>��&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>���پ߱��ʽ�</b><br>���������е��ʽ����ﵽ���ֵ���������������</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="35" name="PStock_Setting(1)" value="<%=PStock_Setting(1)%>"> Ԫ
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>����������Ŀ</b><br>�������й�Ʊ��Ŀ</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="35" name="PStock_Setting(5)" value="<%=PStock_Setting(5)%>"> ��
		</td>
	</tr>	
	<tr>
		<td class="forumRow"><b>���й�˾������С����</b><br>��д���֣�����С��1����50</td>  
		<td class="forumRow">
			&nbsp;<input type=text name="PStock_Setting(2)" value="<%=PStock_Setting(2)%>"> byte
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>���й�˾������󳤶�</b><br>��д���֣�����С��1����50</td>  
		<td class="forumRow">
			&nbsp;<input type=text name="PStock_Setting(3)" value="<%=PStock_Setting(3)%>"> byte 
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>��˾�����󳤶�</b><br>��д���֣�����С��1����255</td>  
		<td class="forumRow">
			&nbsp;<input type=text name="PStock_Setting(4)" value="<%=PStock_Setting(4)%>"> byte 
		</td>
	</tr>

	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=AI_Setting></a>������(AI)���� [<a href="#top">����</a>] [<a href="#bottom">�ײ�</a>]</th>
	</tr>
	<tr>
		<td class="forumRow"><b>�Ƿ���������</b><br>�Ƿ����������¼���ֻ�и���򿪺���������ò���Ч</td> 
		<td class="forumRow">
			<input type=radio name="AI_Setting(0)" value=1 <%if cint(AI_Setting(0))=1 then%>checked<%end if%>>����&nbsp;
			<input type=radio name="AI_Setting(0)" value=0 <%if cint(AI_Setting(0))=0 then%>checked<%end if%>>�ر�&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>AI��������</b><br>�����˷���������Ʊ�����ĸ���<br>0��1֮���С����ֵԽ����Խ��</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="AI_Setting(1)" value="<%=AI_Setting(1)%>">
		</td>
	</tr>
	
	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=KaiHu_Setting></a>���п������� [<a href="#top">����</a>] [<a href="#bottom">�ײ�</a>]</th>
	</tr>
	<tr>
		<td class="forumRow"><b>�Ƿ�������</b></td> 
		<td class="forumRow">
			<input type=radio name="KaiHu_Setting(0)" value=1 <%if cint(KaiHu_Setting(0))=1 then%>checked<%end if%>>��&nbsp;
			<input type=radio name="KaiHu_Setting(0)" value=0 <%if cint(KaiHu_Setting(0))=0 then%>checked<%end if%>>��&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>���ٷ�����</b><br>���Ʊ���ﵽ�����Ŀ�����ڹ��п���<br>��������������Ϊ0</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="KaiHu_Setting(1)" value="<%=KaiHu_Setting(1)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>��������ֵ</b><br>���Ʊ���ﵽ�����Ŀ�����ڹ��п���<br>��������������Ϊ0</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="KaiHu_Setting(2)" value="<%=KaiHu_Setting(2)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>���پ���ֵ</b><br>���Ʊ���ﵽ�����Ŀ�����ڹ��п���<br>��������������Ϊ0</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="KaiHu_Setting(3)" value="<%=KaiHu_Setting(3)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>���ٽ�Ǯֵ</b><br>���Ʊ���ﵽ�����Ŀ�����ڹ��п���<br>��������������Ϊ0</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="KaiHu_Setting(4)" value="<%=KaiHu_Setting(4)%>">
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>ԭʼ�ʱ�</b><br>���͸������ߵ��ʱ�</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="10" name="KaiHu_Setting(6)" value="<%=KaiHu_Setting(6)%>"> Ԫ
		</td>
	</tr>
	
	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=User_Setting></a>�����ʻ����� [<a href="#top">����</a>] [<a href="#bottom">�ײ�</a>]</th>
	</tr>
	<tr>
		<td class="forumRow"><b>�Ƿ�����ӹ�Ʊ�ʻ����</b></td> 
		<td class="forumRow">
			<input type=radio name="User_Setting(0)" value=1 <%if cint(User_Setting(0))=1 then%>checked<%end if%>>��&nbsp;
			<input type=radio name="User_Setting(0)" value=0 <%if cint(User_Setting(0))=0 then%>checked<%end if%>>��&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>�Ƿ������Ǯ���Ʊ�ʻ�</b></td>  
		<td class="forumRow">
			<input type=radio name="User_Setting(1)" value=1 <%if cint(User_Setting(1))=1 then%>checked<%end if%>>��&nbsp;
			<input type=radio name="User_Setting(1)" value=0 <%if cint(User_Setting(1))=0 then%>checked<%end if%>>��&nbsp;		
		</td>
	</tr>
	<tr>
		<td class="forumRow"><b>ȡ������</b><br>����ÿ��ȡ�������<br>��������������Ϊ0</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="User_Setting(2)" value="<%=User_Setting(2)%>"> Ԫ
		</td>
	</tr>

	<tr>
		<td class="forumRow"><b>�������</b><br>����ÿ�δ�������<br>��������������Ϊ0</td>  
		<td class="forumRow">
			&nbsp;<input type=text size="8" name="User_Setting(3)" value="<%=User_Setting(3)%>"> Ԫ
		</td>
	</tr>
	
	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=Custom_Setting></a>�Զ���λ������ [<a href="#top">����</a>] [<a href="#bottom">�ײ�</a>]</th>
	</tr>
	<tr>
		<td class="forumRow" width="40%"><b>�Զ������</b><br>������ʾ�����ݲ�Ҫ̫��<br><font color=red>֧��HTML�﷨</font></td> 
		<td class="forumRow" width="60%">
			<input type=text name="Custom_Setting(0)" size="74" value="<%=Custom_Setting(0)%>">&nbsp;
		</td>
	</tr>
	<tr>
		<td class="forumRow" width="40%"><b>�Զ�������</b><br>������ͼƬ�����ֻ������Ϣ<br>�����ͼƬ��ע��ͼƬ�Ĵ�С(196*52)<br>������ʾ�����ݲ�Ҫ̫��<br><font color=red>֧��HTML�﷨</font></td>   
		<td class="forumRow" width="60%">
			<textarea name="Custom_Setting(1)" cols="75" rows="5"><%=Custom_Setting(1)%></textarea>
		</td>
	</tr>
	
	<tr>
		<th colspan="2" height="25" class="admin" align="left" id="TableTitleLink"><a name=Rule_Setting></a>���й������� [<a href="#top">����</a>] [<a href="#bottom">�ײ�</a>]</th>
	</tr>
	<tr>
		<td class="forumRow" width="40%"><b>��ͣ����</b><br>����Ʊ������ͣ״̬ʱ�Ƿ�������������</td> 
		<td class="forumRow" width="60%">
			<input type=radio name="Trade_Setting(16)" value=0 <%if cint(Trade_Setting(16))=0 then%> checked <%end if%>>��������&nbsp;
			<input type=radio name="Trade_Setting(16)" value=1 <%if cint(Trade_Setting(16))=1 then%> checked <%end if%>>����������&nbsp;
			<input type=radio name="Trade_Setting(16)" value=2 <%if cint(Trade_Setting(16))=2 then%> checked <%end if%>>��������&nbsp;
			<input type=radio name="Trade_Setting(16)" value=3 <%if cint(Trade_Setting(16))=3 then%> checked <%end if%>>������&nbsp;
		</td>
	</tr>
	<tr>
		<td class="forumRow" width="40%"><b>��ͣ����</b><br>����Ʊ���ֵ�ͣ״̬ʱ�Ƿ�������������</td>   
		<td class="forumRow" width="60%">
			<input type=radio name="Trade_Setting(17)" value=0 <%if cint(Trade_Setting(17))=0 then%> checked <%end if%>>��������&nbsp;
			<input type=radio name="Trade_Setting(17)" value=1 <%if cint(Trade_Setting(17))=1 then%> checked <%end if%>>����������&nbsp; 
			<input type=radio name="Trade_Setting(17)" value=2 <%if cint(Trade_Setting(17))=2 then%> checked <%end if%>>��������&nbsp;
			<input type=radio name="Trade_Setting(17)" value=3 <%if cint(Trade_Setting(17))=3 then%> checked <%end if%>>������&nbsp; 
		</td> 
	</tr> 

	<tr>
		<td class="forumRow" colspan="2" align="center" height="10"></td>  
	</tr>	
	<tr>
		<td class="forumRow" colspan="2" align="center"><a name=bottom></a><input type="submit" name="Submit" value="�� ��"></td>  
	</tr>
	<a name=bottom></a>	
	</form>																							
</table>	
<% 
end sub
'-----------------------
sub SaveSettimg()
	dim KaiHu_Setting,PStock_Setting,Gupiao_Setting,Trade_Setting,User_Setting,StopReadme,AI_Setting,Custom_Setting,Rule_Setting
	Gupiao_Setting=request.Form("Gupiao_Setting(0)")&","&request.Form("Gupiao_Setting(1)")&","&request.Form("Gupiao_Setting(2)")&","&request.Form("Gupiao_Setting(3)")&","&request.Form("Gupiao_Setting(4)")&","&request.Form("Gupiao_Setting(5)")&","&request.Form("Gupiao_Setting(6)")&","&request.Form("Gupiao_Setting(7)")&","&request.Form("Gupiao_Setting(8)")&","&request.Form("Gupiao_Setting(9)")&","&request.Form("Gupiao_Setting(10)")&","&request.Form("Gupiao_Setting(11)")&","&request.Form("Gupiao_Setting(12)")
	PStock_Setting=request.Form("PStock_Setting(0)")&"|"&request.Form("PStock_Setting(1)")&"|"&request.Form("PStock_Setting(2)")&"|"&request.Form("PStock_Setting(3)")&"|"&request.Form("PStock_Setting(4)")&"|"&request.Form("PStock_Setting(5)")
	KaiHu_Setting =request.Form("KaiHu_Setting(0)") &"|"&request.Form("KaiHu_Setting(1)") &"|"&request.Form("KaiHu_Setting(2)") &"|"&request.Form("KaiHu_Setting(3)") &"|"&request.Form("KaiHu_Setting(4)") &"|"&request.Form("KaiHu_Setting(5)") &"|"&request.Form("KaiHu_Setting(6)")
	Trade_Setting =request.Form("Trade_Setting(0)") &"||"&request.Form("Trade_Setting(2)")&"|"&request.Form("Trade_Setting(3)")&"|"&request.Form("Trade_Setting(4)")&"|"&request.Form("Trade_Setting(5)")&"|"&request.Form("Trade_Setting(6)")&"|"&request.Form("Trade_Setting(7)")&"|"&request.Form("Trade_Setting(8)")&"|"&request.Form("Trade_Setting(9)")&"|"&request.Form("Trade_Setting(10)")&"|"&request.Form("Trade_Setting(11)")&"|"&request.Form("Trade_Setting(12)") &"|"&request.Form("Trade_Setting(13)")&"|"&request.Form("Trade_Setting(14)")&"|"&request.Form("Trade_Setting(15)")&"|"&request.Form("Trade_Setting(16)")&"|"&request.Form("Trade_Setting(17)")
	User_Setting  =request.Form("User_Setting(0)")&"|"&request.Form("User_Setting(1)")&"|"&request.Form("User_Setting(2)")&"|"&request.Form("User_Setting(3)")
	AI_Setting    =request.Form("AI_Setting(0)")&"|"&request.Form("AI_Setting(1)")
	StopReadme    =iif(request.Form("StopReadme")<>"","StopReadme='"&request.Form("StopReadme")&"'","StopReadme=null")
	Custom_Setting=request.Form("Custom_Setting(0)")&"||"&request.Form("Custom_Setting(1)")
	'Rule_Setting=request.Form("Rule_Setting(0)")&"|"&request.Form("Rule_Setting(1)")
	set rs=conn.execute("update [GupiaoConfig] set Gupiao_Setting='"&Gupiao_Setting&"',PStock_Setting='"&PStock_Setting&"',KaiHu_Setting='"&KaiHu_Setting&"',User_Setting='"&User_Setting&"',Trade_Setting='"&Trade_Setting&"',"&StopReadme&",AI_Setting='"&AI_Setting&"',Custom_Setting='"&Custom_Setting&"'")
	sucmess="<li>��Ʊ��Ϣ���óɹ�"
	call endinfo(1)
end sub
%>