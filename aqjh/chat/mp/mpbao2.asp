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
	sql="SELECT 门派,身份 FROM 用户 where 姓名='" & aqjh_name &"'"
	rs.open sql,conn,1,1
    life1=trim(rs("身份"))
    pai=rs("门派")
    if life1<>"掌门" then
      Response.Write "<script Language=Javascript>alert('提示：你不是掌门,无法代表门派申请门派保护！');location.href = 'javascript:history.go(-1)';</script>"
      rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
        rs.close
       sql="SELECT f FROM p WHERE  a='" & pai & "'and b='"&aqjh_name&"'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
      Response.Write "<script Language=Javascript>alert('提示：你不是掌门,无法代表门派申请门派保护！');location.href = 'javascript:history.go(-1)';</script>"
     rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
	protect1=rs("f")
	   if protect1<>1 then
       Response.Write "<script language=JavaScript>{alert('你们门派现在已经关闭保护,不需要再次关闭');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       conn.execute"update p set f=0 where a='" & pai & "'"
        Response.Write "<script language=JavaScript>{alert('你现在已经脱离官府保护,万事小心.不要过不了多久门派就被攻陷了.');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End

%>