<%'����ս����wWw.51eline.com��
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
if Weekday(date())<>5 or (Hour(time())<20 or Hour(time())>=22) then
Response.Write "<script Language=Javascript>alert('��ʾ����������ս��������ÿ��������20-22����У�����ֻ�����Ų����ʸ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
conn.open connstr
	sql="SELECT ����,���,���� FROM �û� where ����='" & sjjh_name &"'"
	rs.open sql,conn,1,1
    life1=trim(rs("���"))
    pai=rs("����")
    if life1<>"����" then
      Response.Write "<script Language=Javascript>alert('��ʾ���㲻�����ţ��޷�������������ս������');location.href = 'javascript:history.go(-1)';</script>"
      rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
     if rs("����")<10000000 then
     Response.Write "<script language=JavaScript>{alert('������������޷�����ս����������ҲҪ��1ǧ����������');}</script>"
     rs.close
	 set rs=nothing	
	 conn.close
	 set conn=nothing
     Response.End
     end if  
        rs.close
       sql="SELECT f,h FROM p WHERE  a='" & pai & "'and b='"&sjjh_name&"'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
      Response.Write "<script Language=Javascript>alert('��ʾ���㲻������,�޷�������������ս������');location.href = 'javascript:history.go(-1)';</script>"
     rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
	    conn.execute"update p set s=s+100000 where a='" & pai & "'"
	    conn.execute"update �û� set ����=����-10000000 where ����='" & sjjh_name & "'"
        Response.Write "<script language=JavaScript>{alert('��ϲ����Լ�����������ս����10���Լ�������ʧ1ǧ��');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End

%>