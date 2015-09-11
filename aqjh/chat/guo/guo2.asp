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
Response.Write "<script Language=Javascript>alert('提示：增加国家战斗力请于每周一晚上18-22点进行！并且只有君王才有资格增加！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
conn.open connstr
	sql="SELECT 国家,职位,银两,金币,金卡,武功 FROM 用户 where 姓名='" & aqjh_name &"'"
	rs.open sql,conn,1,1
    zhiwei1=trim(rs("职位"))
    guo=rs("国家")
    if zhiwei1<>"君主" then
      Response.Write "<script Language=Javascript>alert('提示：你不是君王,无法代表国家增加战斗力！');location.href = 'javascript:history.go(-1)';</script>"
      rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
     if rs("银两")<50000000 then
     Response.Write "<script language=JavaScript>{alert('你的银两不够无法增加战斗力，至少也要有5千万银两啊！');}</script>"
     rs.close
	 set rs=nothing	
	 conn.close
	 set conn=nothing
     Response.End
     end if  
     if rs("金币")<100 then
     Response.Write "<script language=JavaScript>{alert('你的金币不够无法增加战斗力，至少也要有100个啊！');}</script>"
     rs.close
	 set rs=nothing	
	 conn.close
	 set conn=nothing
     Response.End
     end if  
    if rs("金卡")<2 then
     Response.Write "<script language=JavaScript>{alert('增加国家战斗力收取手续费用2张金卡，你还没有吧');}</script>"
     rs.close
	 set rs=nothing	
	 conn.close
	 set conn=nothing
     Response.End
     end if  
  if rs("武功")<1000 then
     Response.Write "<script language=JavaScript>{alert('增加国家战斗力要扣掉你1000的武功修为，嘿嘿，多练练吧');}</script>"
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
	    conn.execute"update guo set l=l+50000 where g='" & guo & "'"
	    conn.execute"update 用户 set 银两=银两-10000000 金币=金币-100 金卡=金卡-2 武功=武功-1000 where 姓名='" & aqjh_name & "'"
        Response.Write "<script language=JavaScript>{alert('恭喜你帮自己国家增加了战斗力5万,自己银两损失1千万，金币100个，金卡2张，为了损失了1000的武功！.');location.href = 'javascript:history.go(-1)';}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End

%>