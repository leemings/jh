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
      Response.Write "<script Language=Javascript>alert('提示：你不是君王,无法代表国家向官府申请保护！');location.href = 'javascript:history.go(-1)';</script>"
      rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
        rs.close
       sql="SELECT bh,jin FROM guo WHERE  g='" & guo & "'and j='"&aqjh_name&"'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
      Response.Write "<script Language=Javascript>alert('提示：你不是君王,无法代表国家向官府申请保护！');location.href = 'javascript:history.go(-1)';</script>"
     rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
	protect1=rs("bh")
	kujin=rs("jin")
	kujin2=int(kujin*9/10)
	   if protect1=1 then
       Response.Write "<script language=JavaScript>{alert('你们门派现在已经接受官府保护中,不需要再接受保护了');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       conn.execute"update guo set bh=1,jin='"&kujin2&"' where g='" & guo & "'"
        Response.Write "<script language=JavaScript>{alert('你们国家上交了国库十分之一的税,成功的接受官府的保护.');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End

%>