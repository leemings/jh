<%'门派保护♀wWw.happyjh.com♀
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
Response.Write "<script Language=Javascript>alert('提示：门派大战只可以在周五21-22点进行，申请门派保护只有掌门才有资格！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
conn.open connstr
	sql="SELECT 门派,身份 FROM 用户 where 姓名='" & sjjh_name &"'"
	rs.open sql,conn,1,1
    shenfen1=trim(rs("身份"))
    pai=rs("门派")
    if shenfen1<>"掌门" then
      Response.Write "<script Language=Javascript>alert('提示：你不是掌门,无法代表门派申请门派保护！');location.href = 'javascript:history.go(-1)';</script>"
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
      Response.Write "<script Language=Javascript>alert('提示：你不是掌门,无法代表门派申请门派保护！');location.href = 'javascript:history.go(-1)';</script>"
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
       Response.Write "<script language=JavaScript>{alert('你们门派现在已经接受官府保护中，不需要再接受保护了！');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       conn.execute"update p set f=1,h='"&kujin2&"' where a='" & pai & "'"
        Response.Write "<script language=JavaScript>{alert('你们门派上交了库金十分之一的税，成功的接受官府的保护！');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End

%>