<%Option Explicit%><!--#include file="conn.asp"--><%
dim rs
dim formsize,formdata,bncrlf,divider,datastart,dataend,mydata,jiutian,username,id,name

formsize=request.totalbytes
formdata=request.binaryread(formsize)
bncrlf=chrB(13) & chrB(10)
divider=leftB(formdata,clng(instrb(formdata,bncrlf))-1)
datastart=instrb(formdata,bncrlf & bncrlf)+4
dataend=instrb(datastart+1,formdata,divider)-datastart
mydata=midb(formdata,datastart,dataend)

jiutian=Session("jiutian")
username=Session("sjjh_name")

Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "select * from 名妓 where 姓名='"&username&"'",conn,3,2
if rs("photo")>3 then 

response.write "你上传的照片已经超过4张的最大限制。"
response.write "<br><a href=sendphoto.asp>返回</a>"
response.end
end if


rs("photo")=rs("photo")+1
rs.update
id=rs("id")
name=rs("姓名")

%><!--#include file="connpic.asp"--><%

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "pic",conn,3,2
rs.addnew
rs("big").appendchunk mydata
rs("user_id")=id
rs("date")=date
rs("姓名")=name
rs.update

set rs=nothing
set conn=nothing
response.redirect "sendphoto.asp"
%>