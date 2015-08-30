<%dim LocalPic,PicPath
LocalPic=0
if LocalPic=1 then
	PicPath="http://127.0.0.1/eline/bbs/Images/Img_Visual/Show/"
elseif LocalPic=2 then
	PicPath="http://www.shuijingjing.com/images/img_visual/show/"
else
	PicPath="http://qqshow-item.tencent.com/"
end if%>
<%

sub showvisual(uservisual,visualwidth,usersex,userface,userfacewidth,userfaceheight)
 response.write "<table border=0 cellspacing=0 cellpadding=0 width="&visualwidth&">"
 response.write "<tr>"
 response.write "<td align=center visualwidth="&visualwidth&" visualheight="&(visualwidth*226/140)&" ImagePath="""&PicPath&"""   usergender="&usersex&" baseName=""http://127.0.0.1/eline/bbs/images/img_visual/blank.gif"" visual="""&uservisual&"""  style=behavior:url(http://127.0.0.1/eline/bbs/inc/z_show333.htc) localpic="""&localpic&""" userface="""&userface&""" facewidth="&userfacewidth&" faceheight="&userfaceheight&"></td>"
 response.write "</tr>"
 response.write "</table>"
end sub




sub myshow(username,visualwidth,sex)

dim tconn,tconnstr,tdb,trs,tsql
'更改数据库名字，下面改成你自己的数据库
tdb="bbs/data/eline_bbs_6.3.0.asp"
Set tconn = Server.CreateObject("ADODB.Connection")
tconnstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(tdb)
tconn.Open tconnstr
tsql="select * from [user] where username='"&username&"'"
set trs=tconn.execute(tsql)


dim jhconn,jhsql,jhrs

  Set jhconn=Server.CreateObject("ADODB.CONNECTION")
  jhconn.open Application("sjjh_usermdb")
  Set jhrs=Server.CreateObject("ADODB.RecordSet")
  jhsql="select 性别 from 用户 where 姓名='"& action &"'"
  jhrs.open jhsql,jhconn,1,1

if jhrs("性别")="男" then
vsex=1
end if
if jhrs("性别")="女" then
vsex=0
end if




if trs.bof or trs.eof then 
tsql="select * from [user] where username='guest"&vsex&"'"
set trs=tconn.execute(tsql)
elseif trs("visual")="" or isnull(trs("visual")) then
tsql="select * from [user] where username='guest"&vsex&"'"
set trs=tconn.execute(tsql)
end if

call showvisual(trs("visual"),visualwidth,trs("sex"),trs("face"),trs("width"),trs("height"))
set trs=nothing
set tconn=nothing
end sub


%>
<%
'::一线网络www.happyjh.com::%>
<DIV style='padding:0;position:absolute;top:322;left:40;'>
<%
call myshow(action,140,sex)%>
</DIV>