<%@ LANGUAGE=VBScript%>
<% 
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Session("jiutian_name")="" or Session("jiutian_id")="" then response.end
nickname=Session("jiutian_name")
grade=Session("jiutian_grade")
myid=session("jiutian_id")	
If Session("ADMIN") <> True Then Response.Redirect "../error.asp?id=439"
if grade<9 then Response.Redirect "../error.asp?id=439"
Response.Buffer=true
u1=Request.Form("1")
u2=Request.Form("2")
u3=Request.Form("3")
u4=Request.Form("4")
u5=Request.Form("5")
u5=int(u5)
u6=Request.Form("6")
u6=int(u6)
u7=Request.Form("7")
u7=int(u7)
u8=Request.Form("8")
u8=int(u8)
u9=Request.Form("9")
u9=int(u9)
u10=Request.Form("10")
u10=int(u10)
u11=Request.Form("11")
u11=int(u11)
u12=Request.Form("12")
u12=int(u12)
u13=Request.Form("13")
u13=int(u13)
u14=Request.Form("14")
u14=int(u14)
u15=Request.Form("15")
u15=int(u15)
u16=Request.Form("16")
u16=int(u16)
u18=Request.Form("18")
u19=Request.Form("19")
u20=Request.Form("20")
u21=Request.Form("21")
u22=Request.Form("22")
u23=Request.Form("23")
u24=Request.Form("24")
u25=Request.Form("25")
u26=Request.Form("26")
u27=Request.Form("27")
if u14>=11 then response.redirect "manuser.asp"
if u1="" or u2="" or u3="" or u4="" or u5="" or u6="" or  u7="" or u8="" or u9="" or u10="" or  u11="" or u12="" or u13="" or u14="" or u15="" or u16="" then Response.Redirect "../error.asp?id=469"
if instr(u1,"'")<>0 or instr(u1,"|")<>0 or instr(u1," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u2,"'")<>0 or instr(u2,"|")<>0 or instr(u2," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u3,"'")<>0 or instr(u3,"|")<>0 or instr(u3," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u4,"'")<>0 or instr(u4,"|")<>0 or instr(u4," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u5,"'")<>0 or instr(u5,"|")<>0 or instr(u5," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u6,"'")<>0 or instr(u6,"|")<>0 or instr(u6," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u7,"'")<>0 or instr(u7,"|")<>0 or instr(u7," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u8,"'")<>0 or instr(u8,"|")<>0 or instr(u8," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u9,"'")<>0 or instr(u9,"|")<>0 or instr(u9," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u10,"'")<>0 or instr(u10,"|")<>0 or instr(u10," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u11,"'")<>0 or instr(u11,"|")<>0 or instr(u11," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u12,"'")<>0 or instr(u12,"|")<>0 or instr(u12," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u13,"'")<>0 or instr(u13,"|")<>0 or instr(u13," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u14,"'")<>0 or instr(u14,"|")<>0 or instr(u14," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u15,"'")<>0 or instr(u15,"|")<>0 or instr(u15," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u16,"'")<>0 or instr(u16,"|")<>0 or instr(u16," ")<>0 then Response.Redirect "../error.asp?id=470"
if instr(u17,"'")<>0 or instr(u17,"|")<>0 or instr(u17," ")<>0 then Response.Redirect "../error.asp?id=470"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("lbjhhg_connstr")
conn.open connstr
if request.form("submit")="确 定" then
if u14>5  then
if u3<>"六扇门" then 
response.write "此人不是六扇门的人，不可以提升至5级以上"
response.end
end if
else 
if u3="六扇门" then
response.write "此人是六扇门的人，但等级却低于5级"
response.end
end if
end if
conn.execute("update 用户 set 身份='"&u4&"',配偶='"&u2&"',门派='"&u3&"',武功='"&u5&"',内力='"&u6&"',经验='"&u7&"',体力='"&u8&"',攻击='"&u9&"',防御='"&u10&"',银两='"&u11&"',魅力='"&u12&"',道德='"&u13&"',grade='"&u14&"',mvalue='"&u15&"',存款='"&u16&"',状态='"&u17&"' ,师傅='"&u18&"',领钱='"&u19&"',等级='"&u20&"',allvalue='"&u21&"',武器名='"&u22&"' ,防具名='"&u23&"',特武攻='"&u24&"',特武防='"&u25&"',介绍人='"&u26&"',会员='"&u27&"'where 姓名='"&u1&"'")
else
conn.execute("delete * from 用户  where 姓名='"&u1&"'")
end if

response.redirect "../ok.asp?id=700"
%>
