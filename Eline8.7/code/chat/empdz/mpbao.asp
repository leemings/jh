<%'���ɱ�����wWw.happyjh.com��
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
if Weekday(date())<>6 or (Hour(time())<21 or Hour(time())>=22) then
Response.Write "<script Language=Javascript>alert('��ʾ�����ɴ�սֻ����������21-22����У��������ɱ���ֻ�����Ų����ʸ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
conn.open connstr
	sql="SELECT ����,��� FROM �û� where ����='" & sjjh_name &"'"
	rs.open sql,conn,1,1
    shenfen1=trim(rs("���"))
    pai=rs("����")
    if shenfen1<>"����" then
      Response.Write "<script Language=Javascript>alert('��ʾ���㲻������,�޷����������������ɱ�����');location.href = 'javascript:history.go(-1)';</script>"
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
      Response.Write "<script Language=Javascript>alert('��ʾ���㲻������,�޷����������������ɱ�����');location.href = 'javascript:history.go(-1)';</script>"
     rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
	baohu1=rs("f")
	kujin=rs("h")
	kujin2=int(kujin*9/10)
	   if baohu1=1 then
       Response.Write "<script language=JavaScript>{alert('�������������Ѿ����ܹٸ������У�����Ҫ�ٽ��ܱ����ˣ�');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       conn.execute"update p set f=1,h='"&kujin2&"' where a='" & pai & "'"
        Response.Write "<script language=JavaScript>{alert('���������Ͻ��˿��ʮ��֮һ��˰���ɹ��Ľ��ܹٸ��ı�����');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End

%>