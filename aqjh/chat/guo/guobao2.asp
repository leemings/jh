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
	sql="SELECT 国家,职位 FROM 用户 where 姓名='" & aqjh_name &"'"
	rs.open sql,conn,1,1
    zhiwei1=trim(rs("职位"))
    guo=rs("国家")
    if zhiwei1<>"君主" then
      Response.Write "<script Language=Javascript>alert('提示：你不是国家之首,无法向官府申请国家保护！');location.href = 'javascript:history.go(-1)';</script>"
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
      Response.Write "<script Language=Javascript>alert('提示：你不是国家之首,无法向官府申请国家保护！');location.href = 'javascript:history.go(-1)';</script>"
     rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
	protect1=rs("bh")
	   if protect1<>1 then
       Response.Write "<script language=JavaScript>{alert('你们国家现在已经关闭保护,不需要再次关闭');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       conn.execute"update guo set bh=0 where g='" & pai & "'"
        Response.Write "<script language=JavaScript>{alert('你们国家现在已经脱离官府保护,万事小心.国家被攻陷可不是闹着玩的！.');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End

%>