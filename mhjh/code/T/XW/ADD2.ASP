<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
un=session("yx8_mhjh_username")
usergrade=session("yx8_mhjh_usergrade")
if un="" then Response.Redirect "../../error.asp?id=016"
if usergrade<10 then 
Response.Write "<script language=javascript>{alert('对不起，你等级不够10！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
mc=trim(request.form("a"))
lx=trim(request.form("b"))
tu=trim(request.form("c"))
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") or instr(elevalue,"?") or instr(elevalue,"&") or instr(elevalue,"<")  or instr(elevalue,">") or instr(elevalue,"or")  then Response.Redirect "../../error.asp?id=056"
next
if mc="" or lx="" or tu="" then Response.End 
if not(lx="gj" or lx="fs") then Response.End 
if len(mc)>4 then Response.End
function chuser(u)
dim filter,xx,mcenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yushu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
if chuser(mc) then Response.Redirect "../../error.asp?id=524" 
for i=1 to len(mc)
mcchr=mid(mc,i,1)
if asc(mcchr)>0 then Response.Redirect "../../error.asp?id=003"
next
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
        response.write "禁止站外提交！"
        response.end
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"xw")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.open "SELECT 银两,体力,内力,师傅,精力 FROM 用户 where 姓名='"&un&"'",conn
if rst.bof or rst.eof then
	rst.close
	conn.close
	set conn=nothing
	set rst=nothing
	Response.Redirect "../../error.asp?id=016"
	response.end
end if
if rst("银两")<3000000 or rst("精力")<10000 then
	rst.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('需要300万两！精力1万!');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
if rst("师傅")="" or rst("师傅")="无" then
	rst.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('你还没师傅呢，谁教你武功？!');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
if rst("体力")<100000 or rst("内力")<100000 then
	rst.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('需要体力和内力各10万！');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
rst.close
%><!--#include file="data.asp"--><%
sql="select 名称 from mj where 名称='"&mc&"' and 所有者='"&un&"' and 类型='轩'"
set rs=connb.execute(sql)
if not(rs.eof and rs.bof) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connb.close
	set connb=nothing
	Response.Write "<script Language=Javascript>alert('提示：此武器已经存在！');location.href = 'javascript:history.back()';</script>"
	response.end
end if
rs.Close 
      rs.Open "select count(*) as cnumber from mj where  所有者='"&un&"' and 类型='轩'",connb
      if not(rs.eof and rs.bof) then
         cnumber=rs("cnumber")
      end if
if cnumber>=3 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connb.close
	set connb=nothing
	Response.Write "<script language=JavaScript>{alert('要塞武器只能设置三种！');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if   
rs.close
conn.execute "update 用户 set 精力=精力-10000,体力=体力-100000,内力=内力-100000,银两=银两-3000000 where 姓名='"&un&"'"
sqlstr="SELECT * FROM mj "
rs.Open sqlstr,connb,1,2
rs.AddNew
rs("名称")=mc
rs("类型")="轩"
rs("所有者")=session("yx8_mhjh_username")
if lx="gj" then
rs("攻击")=100
rs("防御")=0
rs("状态")=True
else
rs("攻击")=0
rs("防御")=100
rs("状态")=False
end if
rs("速度")=0
rs("等级")=0
rs("特效")="无"
rs("图")=tu
rs("时间")=now()
rs.Update
rs.close
set rs=nothing
connb.close
set connb=nothing
set rst=nothing
conn.close
set conn=nothing
	Response.Write "<script language=JavaScript>{alert('成功获得师傅的招试，你把自己的体力、内力减少10万,精力1万,银两300万送给了师傅，可惜你师傅劳累不堪，一点也没接受到，嘿嘿，好惨，快早点把这老东西超过吧！');location.href = 'xx.asp';}</script>"
%>
