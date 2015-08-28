<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
lastip=Request.ServerVariables("LOCAL_ADDR")
templastip=StrReverse(left(lastip&"9fhb6utime,./",15))
templenlastip=len(lastip)
mmcheck=""
for j=1 to 15
mmcheck=mmcheck+chr(asc(mid(templastip,j,1))-templenlastip+int(j*1.1))
next


username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
mj=request.form("mj")
if not isnumeric(mj) then Response.Redirect  "../error.asp?id=464"
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") or instr(elevalue,"") or instr(elevalue,"?") or instr(elevalue,"#") or instr(elevalue,"&") or instr(elevalue,"or") then Response.Redirect "../error.asp?id=056"
next
if mj="" or mj<1000 then
Response.write "<script language='javascript'>alert ('你太小气了吧，1000两银子也捐不出吗？');location.href='m.asp';</script>"
Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select  银两 from 用户  where 姓名='"&username&"'"
set rs=conn.execute(sql)
if rs("银两")-mj<0 then
Response.write "<script language='javascript'>alert ('您好像钱不够吧，没钱充什么大款?');location.href='m.asp';</script>"
Response.End 
end if
meili=int(mj*.01)
daode=int(mj*.02)
fang=int(mj*.01)
gong=int(mj*.01)
sql="update 用户 set 银两=银两-"&mj&",资质=资质+"&meili&",道德=道德+"&daode&" where 姓名='"&username&"'"
Response.write "<script language='javascript'>alert ('我是魔幻之神，代表大家谢谢你的大义募捐行为！奖励你道德"&daode&"，资质"&meili&"！');location.href='m.asp';</script>"
set rs=conn.Execute(sql)
conn.Close 
set conn=nothing
%>
