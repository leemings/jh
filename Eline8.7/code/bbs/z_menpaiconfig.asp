<%
	dim Meipai_inout_rate,maxtitleid
	dim isZangmen		'�Ƿ�������
	isZangmen=false
	founderr=false
	set rs=conn.execute("select * from [GroupName] where Zangmen='"&membername&"'")
	if not(rs.eof and rs.bof) then isZangmen=true
	
	Meipai_inout_rate=split(conn.execute("select Meipai_inout_rate from [menpai]")(0),"|")
	maxtitleid=conn.execute("select top 1 UserTitleID from [UserTitle] order by UserTitleID desc")(0)
'���� ���� �۳� ��Ǯ������ ���飬���������� �İٷֱ�	1		Meipai_inout_rate(0��3) 		Meipai_inout_rate((menpai_setting(0)-1)*4+0��3))
'���� ħ�� ���� ��Ǯ���۳� ���飬���������� �İٷֱ�	2		Meipai_inout_rate(4��7)			Meipai_inout_rate((menpai_setting(0)-1)*4+0��3))

'�˳� ���� ���� ��Ǯ���۳� ���飬���������� �İٷֱ�	1		Meipai_inout_rate(8��11)		Meipai_inout_rate(menpai_setting(0)*4+5��7)) 		
'�˳� ħ�� �۳� ��Ǯ������ ���飬���������� �İٷֱ�	2		Meipai_inout_rate(12��15) 		Meipai_inout_rate(menpai_setting(0)*4+5��7))

'============================ ����ͨ�ú���============================
'��ȡ�û��ȼ���          
function GetUserTitle(UserClass)
	GetUserTitle=conn.execute("select usertitle from usertitle where userclass="&UserClass&"")(0)
end function
'��ȡ�û��ȼ�����
function GetUserClass(UserTitle)
	GetUserClass=conn.execute("select UserClass from usertitle where UserTitle='"&UserTitle&"'")(0)
end function
'��ȡ��������
function MenpaiAttri(expression)
	if expression=0 or expression="" then
		MenpaiAttri="��������"
	elseif expression=1 then
		MenpaiAttri="����"
	elseif expression=2 then
		MenpaiAttri="ħ��"
	elseif expression=3 then
		MenpaiAttri="<font color=navy>������</font>"
	else
		MenpaiAttri="��������"
	end if
end function
'��ȡ����״̬
function MenpaiState(expression)
	if expression=2 then
		MenpaiState="<font color=red>��ע��</font>"		
	elseif expression=3 then
		MenpaiState="<font color=red>�ȴ����</font>"
	else
		MenpaiState="����"
	end if	
end function
'��ȡ��������ʱ��������ֵ�����İٷ���
function GetMeipai_in_rate(menpaiattri) 
	if cint(menpaiattri)=1 then
		GetMeipai_in_rate="" 
	elseif cint(menpaiattri)=2 then
		GetMeipai_in_rate=""	
	else
		GetMeipai_in_rate="0|0|0|0|0|0"	
	end if	
end function
'�����ַ�������
function strLength(str)
'���ܣ�1��Ӣ����1����λ���ȣ�1��������2����λ����  ������ע��
       ON ERROR RESUME NEXT
       dim WINNT_CHINESE
       WINNT_CHINESE    = (len("��̳")=2)
       if WINNT_CHINESE then
          dim l,t,c
          dim i
          l=len(str)
          t=l
          for i=1 to l
             c=asc(mid(str,i,1))
             if c<0 then c=c+65536
             if c>255 then
                t=t+1
             end if
          next
          strLength=t
       else 
          strLength=len(str)
       end if
       if err.number<>0 then err.clear
end function

'-------------------------------������Ϣ-------------------------------
sub menpai_err() 
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>���ɴ�����Ϣ</th></tr>
	<tr><td width="100%" class=tablebody1><b>��������Ŀ���ԭ��</b><br><br><li>���Ƿ���ϸ�Ķ�������˵�������ɹ��򣬿�������û�е�¼���߲�����ʹ�õ�ǰ���ܵ�Ȩ��
	<font color=<%=Forum_body(8)%>><%=errmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href="javascript:history.go(-1)">������һҳ</a></td></tr> 
</table>
<%
end sub

'-------------------------------�ɹ���Ϣ-------------------------------
sub menpai_suc() 
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>���ɲ����ɹ�</th></tr>
	<tr><td width="100%" class=tablebody1><b>�����ɹ�</b><br><font color=navy><%=sucmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href=<%=Request.ServerVariables("HTTP_REFERER")%>>������һҳ</a></td></tr>
</table>
<%
end sub
%>