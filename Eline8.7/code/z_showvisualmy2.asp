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
 response.write "<td align=left visualwidth="&visualwidth&" visualheight="&(visualwidth*226/140)&" ImagePath="""&PicPath&"""   usergender="&usersex&" baseName=""http://127.0.0.1/eline/bbs/images/img_visual/blank.gif"" visual="""&uservisual&"""  style=behavior:url(http://127.0.0.1/eline/bbs/inc/z_show333.htc) localpic="""&localpic&""" userface="""&userface&""" facewidth="&userfacewidth&" faceheight="&userfaceheight&"></td>"
 response.write "</tr>"
 response.write "</table>"
end sub




sub myshow(username,visualwidth,sex)

dim tconn,tconnstr,tdb,trs,tsql
'�������ݿ����֣�����ĳ����Լ������ݿ�
tdb="/eline/bbs/data/eline_bbs_6.3.0.asp"
Set tconn = Server.CreateObject("ADODB.Connection")
tconnstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(tdb)
tconn.Open tconnstr
tsql="select * from [user] where username='"&username&"'"
set trs=tconn.execute(tsql)


dim jhconn,jhsql,jhrs

  Set jhconn=Server.CreateObject("ADODB.CONNECTION")
  jhconn.open Application("sjjh_usermdb")
  Set jhrs=Server.CreateObject("ADODB.RecordSet")
  jhsql="select �Ա� from �û� where ����='"& sjjh_name &"'"
  jhrs.open jhsql,jhconn,1,1

if jhrs("�Ա�")="��" then
vsex=1
end if
if jhrs("�Ա�")="Ů" then
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
'::һ������www.51eline.com::%>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#e1f0db" height="220" align="center" bordercolorlight="#a3dc98" bordercolordark="#333333">
  <tr> 
        <td width="100%" height="100%">
<%call myshow(session("sjjh_name"),135,sex)%>
</td>
      </tr>
    </table>
