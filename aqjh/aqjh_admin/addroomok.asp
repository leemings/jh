<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
name=trim(Request.Form("name"))
zdrs=abs(trim(Request.Form("zdrs")))
fzxz=abs(trim(Request.Form("fzxz")))
sjxz=abs(trim(Request.Form("sjxz")))
bh=abs(trim(Request.Form("bh")))
kp=abs(trim(Request.Form("kp")))
db=abs(trim(Request.Form("db")))
pd=abs(trim(Request.Form("pd")))
djsj=abs(trim(Request.Form("djsj")))
xz=abs(Request.Form("xz"))
xzsm=trim(Request.Form("xzsm"))
bds=trim(Request.Form("bds"))
jrht=trim(Request.Form("jrht"))
if len(jrht)<4 or len(jrht)>150 then
	Response.Write "<script Language=Javascript>alert('提示：今日话题为4-150个字符!');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if zdrs<2 or zdrs>500 then
	Response.Write "<script Language=Javascript>alert('提示：最大人数为2-500');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if fzxz<>0 and fzxz<>1 then
	Response.Write "<script Language=Javascript>alert('提示：发招选择0为允许，1为禁止');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if djsj<10 or djsj>60 then
	Response.Write "<script Language=Javascript>alert('提示：打架时间设置10-60分');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if pd<>0 and pd<>1 then
	Response.Write "<script Language=Javascript>alert('提示：泡点选择0为允许，1为禁止');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if sjxz<>0 and sjxz<>1 then
	Response.Write "<script Language=Javascript>alert('提示：事件限制选择0为允许，1为禁止');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if bh<>0 and bh<>1 then
	Response.Write "<script Language=Javascript>alert('提示：保护限制选择0为允许，1为禁止');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if kp<>0 and kp<>1 then
	Response.Write "<script Language=Javascript>alert('提示：卡片限制选择0为允许，1为禁止');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if db<>0 and db<>1 then
	Response.Write "<script Language=Javascript>alert('提示：赌博限制选择0为允许，1为禁止');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if xz<>0 and xz<>1 then
	Response.Write "<script Language=Javascript>alert('提示：限制进入选择0为允许，1为禁止');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
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
rs("l")=djsj
rs("m")=pd
rs.Update
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','[新增房间：]"&name&"的设置！')"
conn.close
set conn=nothing
Response.Redirect "../ok.asp?id=700"
%>