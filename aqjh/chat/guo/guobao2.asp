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
conn.open connstr
	sql="SELECT ����,ְλ FROM �û� where ����='" & aqjh_name &"'"
	rs.open sql,conn,1,1
    zhiwei1=trim(rs("ְλ"))
    guo=rs("����")
    if zhiwei1<>"����" then
      Response.Write "<script Language=Javascript>alert('��ʾ���㲻�ǹ���֮��,�޷���ٸ�������ұ�����');location.href = 'javascript:history.go(-1)';</script>"
      rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
        rs.close
       sql="SELECT bh FROM guo WHERE  g='" & guo & "'and j='"&aqjh_name&"'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
      Response.Write "<script Language=Javascript>alert('��ʾ���㲻�ǹ���֮��,�޷���ٸ�������ұ�����');location.href = 'javascript:history.go(-1)';</script>"
     rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
	protect1=rs("bh")
	   if protect1<>1 then
       Response.Write "<script language=JavaScript>{alert('���ǹ��������Ѿ��رձ���,����Ҫ�ٴιر�');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       conn.execute"update guo set bh=0 where g='" & pai & "'"
        Response.Write "<script language=JavaScript>{alert('���ǹ��������Ѿ�����ٸ�����,����С��.���ұ����ݿɲ���������ģ�.');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End

%>