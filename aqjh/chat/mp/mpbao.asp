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
	sql="SELECT ����,��� FROM �û� where ����='" & aqjh_name &"'"
	rs.open sql,conn,1,1
    life1=trim(rs("���"))
    pai=rs("����")
    if life1<>"����" then
      Response.Write "<script Language=Javascript>alert('��ʾ���㲻������,�޷����������������ɱ�����');location.href = 'javascript:history.go(-1)';</script>"
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
      Response.Write "<script Language=Javascript>alert('��ʾ���㲻������,�޷����������������ɱ�����');location.href = 'javascript:history.go(-1)';</script>"
     rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
	protect1=rs("f")
	kujin=rs("h")
	kujin2=int(kujin*9/10)
	   if protect1=1 then
       Response.Write "<script language=JavaScript>{alert('�������������Ѿ����ܹٸ�������,����Ҫ�ٽ��ܱ�����');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       conn.execute"update p set f=1,h='"&kujin2&"' where a='" & pai & "'"
        Response.Write "<script language=JavaScript>{alert('���������Ͻ��˿��ʮ��֮һ��˰,�ɹ��Ľ��ܹٸ��ı���.');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End

%>