<%'���ɱ�����wWw.51eline.com��
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr
	sql="SELECT ����,��� FROM �û� where ����='" & sjjh_name &"'"
	rs.open sql,conn,1,1
    shenfen1=trim(rs("���"))
    pai=rs("����")
    if shenfen1<>"����" then
      Response.Write "<script Language=Javascript>alert('��ʾ���㲻�����ţ��޷����������������ɹرձ�����');location.href = 'javascript:history.go(-1)';</script>"
      rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
        rs.close
       sql="SELECT f FROM p WHERE  a='" & pai & "'and b='"&sjjh_name&"'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
      Response.Write "<script Language=Javascript>alert('��ʾ���㲻�����ţ��޷����������������ɹرձ�����');location.href = 'javascript:history.go(-1)';</script>"
     rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
	baohu1=rs("f")

	   if baohu1<>1 then
       Response.Write "<script language=JavaScript>{alert('�������������Ѿ��رձ���������Ҫ�ٴιرգ�');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       conn.execute"update p set f=0 where a='" & pai & "'"
        Response.Write "<script language=JavaScript>{alert('�������Ѿ�����ٸ�����������С�ģ���Ҫ�����˶�����ɾͱ������ˣ�');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End

%>