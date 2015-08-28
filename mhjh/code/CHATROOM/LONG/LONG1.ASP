<%@ LANGUAGE=VBScript%>
<% 
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
id=request("id")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 or InStr(id,"&")<>0 then Response.Redirect "../../error.asp?id=574"
%><!--#include file="data.asp"--><% 	
sql="SELECT 物品名,精力,攻击,会员,说明,银两 FROM 商店 where 类型='龙' and ID=" & id
rs.open sql,connb,1,1
if rs.eof or rs.bof then
    Response.Write "<script Language=Javascript>alert('没有这样的神龙，是不是在作弊呀!');window.close();</script>"
    rs.close
	set rs=nothing	
	connb.close
	set connb=nothing
    response.end
end if
       wu=rs("物品名")
       gj=rs("攻击")
	   lx=rs("会员")
       dian=int(rs("精力"))
       allgj=rs("银两")
       sm=rs("说明")
rs.close
sql= "SELECT * FROM myanimal where username='"&username&"' and lei='"&wu&"'"   
rs.open sql,connb,1,1
if not (rs.eof or rs.bof) then
    Response.Write "<script Language=Javascript>alert('你已经有这种神龙了，不能再买!');window.close();</script>"
    rs.close
	set rs=nothing	
	connb.close
	set connb=nothing
    response.end
end if
rs.close
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="select 会员,银两,精力 from 用户 where 姓名='"&username&"'"
rs.open sql,conn,1,1
       hy=rs("会员")
       yl=int(rs("银两"))
       yq=int(rs("精力"))
   if lx="是" and hy<>"True" then
       Response.Write "<script Language=Javascript>alert('你不是会员，不能购买这种类型的神龙，此为会员使用');window.close();</script>"
      rs.close
	  set rs=nothing	
	  connb.close
	  set connb=nothing
      response.end
   end if
   if allgj>yl then
      Response.Write "<script Language=Javascript>alert('购买不成功，你身上银两不够呀');window.close();</script>"
      rs.close
	  set rs=nothing
	  connb.close
	  set connb=nothing
      response.end
   end if
   if yq<dian then
      Response.Write "<script Language=Javascript>alert('购买不成功，你身上没多精力呀');window.close();</script>"
      rs.close
	  set rs=nothing
	  connb.close
	  set connb=nothing
      response.end
   end if
         conn.execute"update 用户 set 精力=精力-" & dian & ",银两=银两-"& allgj &"  where 姓名='" & username & "'"
         connb.Execute "insert into myanimal(animalname,username,attack,lei,allattack,sm)values('"&wu&"','"&username&"',"&gj&",'"&wu&"',"&allgj/10&",'"&sm&"')"
	     Response.Write "<script Language=Javascript>alert('购买成功。');location.href='long.asp';</script>"
         rs.close
	     set rs=nothing
	     connb.close
	     set connb=nothing
	     conn.close
	     set conn=nothing
         response.end
%>

