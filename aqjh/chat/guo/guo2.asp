<%
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
if Weekday(date())<>2 or (Hour(time())<18 or Hour(time())>=22) then
Response.Write "<script Language=Javascript>alert('��ʾ�����ӹ���ս��������ÿ��һ����18-22����У�����ֻ�о��������ʸ����ӣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
conn.open connstr
	sql="SELECT ����,ְλ,����,���,��,�书 FROM �û� where ����='" & aqjh_name &"'"
	rs.open sql,conn,1,1
    zhiwei1=trim(rs("ְλ"))
    guo=rs("����")
    if zhiwei1<>"����" then
      Response.Write "<script Language=Javascript>alert('��ʾ���㲻�Ǿ���,�޷������������ս������');location.href = 'javascript:history.go(-1)';</script>"
      rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
     if rs("����")<50000000 then
     Response.Write "<script language=JavaScript>{alert('������������޷�����ս����������ҲҪ��5ǧ����������');}</script>"
     rs.close
	 set rs=nothing	
	 conn.close
	 set conn=nothing
     Response.End
     end if  
     if rs("���")<100 then
     Response.Write "<script language=JavaScript>{alert('��Ľ�Ҳ����޷�����ս����������ҲҪ��100������');}</script>"
     rs.close
	 set rs=nothing	
	 conn.close
	 set conn=nothing
     Response.End
     end if  
    if rs("��")<2 then
     Response.Write "<script language=JavaScript>{alert('���ӹ���ս������ȡ��������2�Ž𿨣��㻹û�а�');}</script>"
     rs.close
	 set rs=nothing	
	 conn.close
	 set conn=nothing
     Response.End
     end if  
  if rs("�书")<1000 then
     Response.Write "<script language=JavaScript>{alert('���ӹ���ս����Ҫ�۵���1000���书��Ϊ���ٺ٣���������');}</script>"
     rs.close
	 set rs=nothing	
	 conn.close
	 set conn=nothing
     Response.End
     end if  
        rs.close
       sql="SELECT f,h FROM p WHERE  a='" & pai & "'and b='"&aqjh_name&"'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
      Response.Write "<script Language=Javascript>alert('��ʾ���㲻������,�޷�������������ս������');location.href = 'javascript:history.go(-1)';</script>"
     rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
	    conn.execute"update guo set l=l+50000 where g='" & guo & "'"
	    conn.execute"update �û� set ����=����-10000000 ���=���-100 ��=��-2 �书=�书-1000 where ����='" & aqjh_name & "'"
        Response.Write "<script language=JavaScript>{alert('��ϲ����Լ�����������ս����5��,�Լ�������ʧ1ǧ�򣬽��100������2�ţ�Ϊ����ʧ��1000���书��.');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End

%>