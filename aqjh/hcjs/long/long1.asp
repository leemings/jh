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

if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../error.asp?id=120"

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
	
sql="SELECT 物品名,内力,体力,攻击,状态,说明,银两无 FROM 物品买 where 类型='龙' and ID=" & id
rs.open sql,conn,1,1

if rs.eof or rs.bof then
       Response.Write "<script Language=Javascript>alert('没有这样的神龙，是不是在作弊呀!');window.close();</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       response.end
end if
       wu=rs("物品名")
       gj=rs("攻击")
	lx=rs("状态")
       jin=int(rs("内力"))
       dian=int(rs("体力"))
       allgj=rs("银两无")
       sm=rs("说明")
rs.close

sql="select * from 用户 where id=" & myid
rs.open sql,conn,1,1
       hy=rs("会员等级")
       jinyan=int(rs("allvalue"))
       dianjuan=int(rs("金币"))

   if lx="是" and hy<2 then
       Response.Write "<script Language=Javascript>alert('你不是2级以上会员，不能购买这种类型的神龙');window.close();</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       response.end
   end if

      if jin>jinyan then
                 Response.Write "<script Language=Javascript>alert('购买不成功，你没这么经验呀');window.close();</script>"
                 rs.close
	          set rs=nothing
	          conn.close
	          set conn=nothing
                 response.end
       end if
	

       if dian>dianjuan then
                 Response.Write "<script Language=Javascript>alert('购买不成功，你身上没这么金币呀');window.close();</script>"
                 rs.close
	          set rs=nothing
	          conn.close
	          set conn=nothing
                 response.end
       end if
   
         conn.execute"update 用户 set allvalue=allvalue-" & jin & ",金币=金币-"& dian &"  where 姓名='" & my & "'"
         conn.Execute "insert into myanimal(animalname,username,attack,lei,allattack,sm)values('"&wu&"','"&my&"',"&gj&",'"&wu&"',"&allgj&",'"&sm&"')"
	   Response.Write "<script Language=Javascript>alert('购买成功。');window.close();</script>"
                 rs.close
	          set rs=nothing
	          conn.close
	          set conn=nothing
                 response.end
                  
	%>


</body>
</html>
