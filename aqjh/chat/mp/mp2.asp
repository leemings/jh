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
Response.Write "<script Language=Javascript>alert('提示：增加门派战斗力请于每周一晚上18-22点进行！并且只有掌门才有资格增加！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
conn.open connstr
	sql="SELECT 门派,身份,银两 FROM 用户 where 姓名='" & aqjh_name &"'"
	rs.open sql,conn,1,1
    life1=trim(rs("身份"))
    pai=rs("门派")
    if life1<>"掌门" then
      Response.Write "<script Language=Javascript>alert('提示：你不是掌门,无法代表门派增加战斗力！');location.href = 'javascript:history.go(-1)';</script>"
      rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
     if rs("银两")<10000000 then
     Response.Write "<script language=JavaScript>{alert('你的银两不够无法增加战斗力，至少也要有1千万银两啊！');}</script>"
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
      Response.Write "<script Language=Javascript>alert('提示：你不是掌门,无法代表门派增加战斗力！');location.href = 'javascript:history.go(-1)';</script>"
     rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
	    conn.execute"update p set s=s+100000 where a='" & pai & "'"
	    conn.execute"update 用户 set 银两=银两-10000000 where 姓名='" & aqjh_name & "'"
        Response.Write "<script language=JavaScript>{alert('恭喜你帮自己门派增加了战斗力10万,自己银两损失1千万.');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End

%>