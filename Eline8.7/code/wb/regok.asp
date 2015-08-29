<%@ LANGUAGE=VBScript codepage ="936" %><%

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "error.asp?id=440"


%>
<!--#include file="dbpath.asp"-->
<!--#include file="char.inc"-->
<%
city=request.form("city")
barname=request.form("barname")
typical=request.form("typical")
number=request.form("number")
add=request.form("add")
people=request.form("people")
pwd=request.form("pwd")
pwd2=request.form("pwd2")
phone=request.form("phone")
email=request.form("email")
zip=request.form("zip")
ip=request.form("ip")
homepage=request.form("homepage")
intros=request.form("intros")


'**********检查是否填写了所有项，如果不是侧自动返回申请页面
if city="" or barname="" or typical="" or add="" or pwd="" or people="" or intros="" then
errmsg=errmsg & "请填写完整网吧信息！\n"
end if
if pwd<>pwd2 then
errmsg=errmsg & "两次输入的密码不相同！\n"
end if

dim rsc,errmsg
Set rsc = Conn.Execute("select * from bar where barname = '" & barname & "'")
if not rsc.eof then errmsg=errmsg & "网吧名称已被注册，请改名！\n"
if sjjh_grade<>10 then
errmsg="你不是管理员你想作什么？！"
end if
if errmsg<>"" then
    Conn.Close
    Set conn = nothing
    Set rsc = nothing
    response.write("<script>alert('" & errmsg & "');history.go(-1)</script>")
    response.end
end if
'**********检查结束**********

set rs=server.createobject("adodb.recordset")
sql="select * from bar where (id is null)" 

rs.open sql,conn,1,3
rs.addnew
rs("city")=htmlencode(city)
rs("barname")=htmlencode(barname)
rs("typical")=htmlencode(typical)
rs("number")=htmlencode(number)
rs("add")=htmlencode(add)
rs("people")=htmlencode(people)
rs("pwd")=pwd
rs("phone")=htmlencode(phone)
rs("email")=htmlencode(email)
rs("zip")=htmlencode(zip)
rs("date")=htmlencode(date())
rs("ip")=htmlencode(ip)
rs("homepage")=htmlencode(homepage)
rs("intros")=htmlencode(intros)

rs.update
rs.close
conn.close
set conn=nothing
Response.Redirect "index.asp?type=regok"%>