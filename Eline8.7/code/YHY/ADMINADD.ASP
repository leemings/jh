<%Option Explicit%><!--#include file="conn.asp"--><%
dim rs
dim formsize,formdata,bncrlf,divider,datastart,dataend,mydata,jiutian,username,id,name,user_id
user_id=request("id")

formsize=request.totalbytes
formdata=request.binaryread(formsize)
bncrlf=chrB(13) & chrB(10)
divider=leftB(formdata,clng(instrb(formdata,bncrlf))-1)
datastart=instrb(formdata,bncrlf & bncrlf)+4
dataend=instrb(datastart+1,formdata,divider)-datastart
mydata=midb(formdata,datastart,dataend)

Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "select * from УћМЫ where id=" & user_id,conn,3,2

rs("photo")=rs("photo")+1
rs.update
id=rs("id")
name=rs("аеУћ")

%><!--#include file="connpic.asp"--><%

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "pic",conn,3,2
rs.addnew
rs("big").appendchunk mydata
rs("user_id")=user_id
rs("date")=date
rs("аеУћ")=name
rs.update

set rs=nothing
set conn=nothing
response.redirect "admin.asp?id=" & user_id

%>