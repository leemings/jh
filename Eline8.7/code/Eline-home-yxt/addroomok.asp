<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
name=trim(Request.Form("name"))
zdrs=abs(trim(Request.Form("zdrs")))
fzxz=abs(trim(Request.Form("fzxz")))
sjxz=abs(trim(Request.Form("sjxz")))
bh=abs(trim(Request.Form("bh")))
kp=abs(trim(Request.Form("kp")))
db=abs(trim(Request.Form("db")))
xz=abs(Request.Form("xz"))
xzsm=trim(Request.Form("xzsm"))
bds=trim(Request.Form("bds"))
jrht=trim(Request.Form("jrht"))
if len(jrht)<4 or len(jrht)>150 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����ջ���Ϊ4-150���ַ�!');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if zdrs<2 or zdrs>500 then
	Response.Write "<script Language=Javascript>alert('��ʾ���������Ϊ2-500');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if fzxz<>0 and fzxz<>1 then
	Response.Write "<script Language=Javascript>alert('��ʾ������ѡ��0Ϊ����1Ϊ��ֹ');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if sjxz<>0 and sjxz<>1 then
	Response.Write "<script Language=Javascript>alert('��ʾ���¼�����ѡ��0Ϊ����1Ϊ��ֹ');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if bh<>0 and bh<>1 then
	Response.Write "<script Language=Javascript>alert('��ʾ����������ѡ��0Ϊ����1Ϊ��ֹ');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if kp<>0 and kp<>1 then
	Response.Write "<script Language=Javascript>alert('��ʾ����Ƭ����ѡ��0Ϊ����1Ϊ��ֹ');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if db<>0 and db<>1 then
	Response.Write "<script Language=Javascript>alert('��ʾ���Ĳ�����ѡ��0Ϊ����1Ϊ��ֹ');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if xz<>0 and xz<>1 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����ƽ���ѡ��0Ϊ����1Ϊ��ֹ');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sqlstr="SELECT * FROM r "
rs.Open sqlstr,conn,1,2
rs.AddNew
rs("a")=name
rs("b")=zdrs
rs("c")=xz
rs("d")=xzsm
rs("e")=bds
rs("f")=fzxz
rs("g")=sjxz
rs("h")=bh
rs("i")=kp
rs("j")=db
rs("k")=jrht
rs.Update
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','[�������䣺]"&name&"�����ã�')"
'conn.execute "insert into r(a,b,c,d,e,f,g,h,i,j) values 
'('"&name&"',"&zdrs&","&xz&",'"&xzsm&"','"&bds&"',"&fzxz&","&sjxz&","&bh&","&kp&","&db&")"
'conn.Execute "update r set a='"& name &"',c="& xz &",d='"& xzsm &"',e='"& bds &"',b="& zdrs &",f="& fzxz &",g="& sjxz &",h="& bh &",i="&kp&",j="&db&" where id="&ID
conn.close
set conn=nothing
Response.Redirect "../ok.asp?id=700"
%>