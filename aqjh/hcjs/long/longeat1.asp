<%@ LANGUAGE=VBScript%>
<!--#include file="check_user.asp"-->

<% 
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

if Session("aqjh_name")="" or Session("aqjh_id")="" then response.end
my=Session("aqjh_name")
grade=Session("aqjh_grade")
myid=session("aqjh_id")
id=request("id")

if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=120"

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
	sql="SELECT * FROM 物品买 where 类型='龙粮' and ID=" & id
	rs.open sql,conn,1,1
if rs.eof or rs.bof then
       Response.Write "<script Language=Javascript>alert('没有这样的物品，是不是在作弊呀!');window.close();</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       response.end
end if
       wu=rs("物品名")
	yin=rs("银两")
       yinw=rs("银两无")
	lx=rs("类型")
       gj=rs("攻击")
       fy=rs("防御")
	zt=rs("状态")
       nl=rs("内力")
       tl=rs("体力")
       sm=rs("说明")
rs.close

sql="select 金币 from 用户 where id=" & myid
rs.open sql,conn,1,1
       jinyan=int(rs("金币"))
	if yin>jinyan then
                 Response.Write "<script Language=Javascript>alert('购买不成功，你没这么多金币呀');window.close();</script>"
                 rs.close
	          set rs=nothing
	          conn.close
	          set conn=nothing
                 response.end
       end if
             
                    conn.execute"update 用户 set 金币=金币-" & yin & " where 姓名='" & my & "'"
			rs.close
                    sql="select 数量 from 物品 where 物品名='" & wu & "' and 拥有者='" & my & "'"
			rs.open sql,conn,1,1
                    
			if rs.eof or rs.bof then
			   conn.execute"insert into 物品(物品名,拥有者,类型,体力,银两) values ('"&wu&"','"&my&"','"&lx&"','"&tl&"','"&yin&"')"
                       conn.execute"update 物品 set 数量=1 where 物品名='" & wu & "' and 拥有者='" & my & "'"
                       Response.Write "<script Language=Javascript>alert('购买" & wu & "成功!');window.close();</script>"

  rs.close
	                 set rs=nothing
	                 conn.close
	                 set conn=nothing

			   Response.Redirect "longeat.asp"
		        end if
                      shu=rs("数量")
                    if shu>10 then 
                       Response.Write "<script Language=Javascript>alert('你身上的物品太多，不能再买!');window.close();</script>"
                       rs.close
	                set rs=nothing
	                conn.close
	                set conn=nothing
                       response.end
                     else	
			   conn.execute"update 物品 set 数量=数量+1  where 物品名='" & wu & "' and 拥有者='" & my & "'"
			    Response.Write "<script Language=Javascript>alert('购买" & wu & "成功!');window.close();</script>"

		         rs.close
	                set rs=nothing
	                conn.close
	                set conn=nothing
			  Response.Redirect "longeat.asp"
                     end if
	%>


</body>
</html>
