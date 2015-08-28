<%@ LANGUAGE=VBScript%>
<% 
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
ll=request("mz")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 or InStr(id,"&")<>0 then Response.Redirect "../../error.asp?id=524"
%><!--#include file="data.asp"--><% 	
sqlstr="select 物品名,银两,类型,攻击,防御,精力,说明 from 商店 where id="&ll&" and 类型='龙粮'"
rs.open sqlstr,connb,1,2
if rs.eof or rs.bof then
    Response.Write "<script Language=Javascript>alert('没有这样的物品，是不是在作弊呀!');window.close();</script>"
    rs.close
	set rs=nothing	
	connb.close
	set connb=nothing
    response.end
end if
       wu=rs("物品名")
	   yin=rs("银两")
	   lx=rs("类型")
       gj=rs("攻击")
       fy=rs("防御")
       nl=int(rs("精力"))
       sm=rs("说明")
rs.close
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="select 银两,精力 from 用户 where 姓名='"&username&"'"
rs.open sql,conn,1,1
       jinyan=int(rs("精力"))
       yl=int(rs("银两"))
	   if yin>yl then
              Response.Write "<script Language=Javascript>alert('购买不成功，你身上的钱不够呀');window.close();</script>"
              rs.close
	          set rs=nothing
	          connb.close
	          set connb=nothing
	          conn.close
	          set conn=nothing
              response.end
       end if
       	   if jinyan<nl then
              Response.Write "<script Language=Javascript>alert('购买不成功，你身上的精力不够呀');window.close();</script>"
              rs.close
	          set rs=nothing
	          connb.close
	          set connb=nothing
	          conn.close
	          set conn=nothing
              response.end
       end if
            conn.execute"update 用户 set 精力=精力-" & nl & ",银两=银两-"& yin &" where 姓名='" & username & "'"
			rs.close
            sql="select 数量 from 物品 where 名称='" & wu & "' and 所有者='"&username&"'"
			rs.open sql,conn,1,2     
			if rs.eof or rs.bof then
			         conn.execute"insert into 物品(名称,所有者,属性,体力,寄售,装备,特效,价格,数量) values ('"&wu&"','"&username&"','"&lx&"','"&nl&"',False,False,'无',0,1)"
                     rs.close
	                 set rs=nothing
	                 connb.close
	                 set connb=nothing
	                 conn.close
	                 set conn=nothing
			          Response.Write "<script Language=Javascript>alert('购买成功。');location.href='longeat.asp';</script>"
		   
             else
			     conn.execute"update 物品 set 数量=数量+1  where 名称='"&wu&"' and 所有者='"&username& "'"
		       rs.close
	           set rs=nothing
	           connb.close
	           set connb=nothing
	           conn.close
	           set conn=nothing
	           Response.Write "<script Language=Javascript>alert('购买成功。');location.href='longeat.asp';</script>"
            end if
	%>


