<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_grade<>10 then Response.Redirect "../error.asp?id=439"
if session("aqjh_adminok")<>true then response.end
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
act=sqlstr(request("act"))
Select Case act
case "addnpc"
	call zzz()
	npcname=sqlstr(request.form("npcname"))
	npcsex=sqlstr(request.form("npcsex"))
	npcvalue=sqlstr(request.form("npcvalue"))
	npcimg=sqlstr(request.form("npcimg"))
	npcmoney=sqlstr(request.form("npcmoney"))
	npcgj=sqlstr(request.form("npcgj"))
	npcfy=sqlstr(request.form("npcfy"))
	npcwg=sqlstr(request.form("npcwg"))
	npctl=sqlstr(request.form("npctl"))
	npcnl=sqlstr(request.form("npcnl"))
	npcdie=sqlstr(request.form("npcdie"))
	npcgjl=sqlstr(request.form("npcgjl"))
	npccxl=sqlstr(request.form("npccxl"))
	npcwp=sqlstr(request.form("npcwp"))
	npcwplx=sqlstr(request.form("npcwplx"))
	npcccc=sqlstr(request.form("npcccc"))
	rs.open "Select * from npc Where n����='" & npcname & "'",conn,1,1
	if not rs.eof then
		rs.close
		call smsg("[" & npcname & "]��npc�Ѿ����ڣ�","-1")
	end if
	rs.close
	rs.Open "SELECT * FROM npc ",conn,1,3
	rs.AddNew
        rs("n����")=npcname
	rs("n�Ա�")=npcsex
	rs("n����")=npcvalue
	rs("nͼƬ")=npcimg
	rs("n����")=npcmoney
	rs("n����")=npcgj
	rs("n����")=npcfy
	rs("n�书")=npcwg
	rs("n����")=npctl
	rs("n����")=npcnl
	rs("n������")=npcdie
	rs("n������")=npcgjl
	rs("n������")=npccxl
	rs("n��Ʒ")=npcwp
	rs("n��Ʒ����")=npcwplx
	rs("n������")=npcccc
	rs.Update
	conn.execute "insert into l(a,b,c,d,e) values (now(),'" & aqjh_name & "','" & Request.ServerVariables("REMOTE_ADDR") & "','�����¼','����NPC:" & npcname & "')"
	rs.close
	call smsg("��ϲ,npc��ӳɹ�!","npclist.asp")

case "delnpc"
	call zzz()
	npcname=sqlstr(request("npcname"))
	conn.Execute "delete * from npc where n����='" & npcname &"'"
	conn.execute "insert into l(a,b,c,d,e) values (now(),'" & aqjh_name & "','" & Request.ServerVariables("REMOTE_ADDR") & "','�����¼','ɾ��npc {id}:[" & id & "]')"
	call smsg("ɾ���ɹ�!","npclist.asp")
case "delczc"
	call zzz()
	sn=sqlstr(request("sn"))
	conn.Execute "delete * from czc where sn='" & sn &"'"
	conn.execute "insert into l(a,b,c,d,e) values (now(),'" & aqjh_name & "','" & Request.ServerVariables("REMOTE_ADDR") & "','�����¼','ɾ����ֵ��SN:[" & sn & "]')"
	call smsg("ɾ���ɹ�!","czclist.asp")
case "delallczc"
	call zzz()
	conn.Execute "delete * from czc where cz_use=1"
	conn.execute "insert into l(a,b,c,d,e) values (now(),'" & aqjh_name & "','" & Request.ServerVariables("REMOTE_ADDR") & "','�����¼','ɾ�������ù���ֵ���ɹ�!')"
	call smsg("ɾ���ɹ�!","czclist.asp")
case "modiczc"
	call zzz()
	id=clng(request("id"))
	pass=request.form("pass")
	cz_use=sqlstr(request.form("cz_use"))
	sm=sqlstr(request.form("sm"))
	'��������!
	rs.Open "SELECT * FROM czc where id=" & id,conn,1,3
	if len(pass)=6 then rs("pass")=md5(pass)
	if cz_use=1 then
		rs("cz_use")=1
	else
		rs("cz_use")=0
	end if
	rs("remark")=sm
	rs.Update
	conn.execute "insert into l(a,b,c,d,e) values (now(),'" & aqjh_name & "','" & Request.ServerVariables("REMOTE_ADDR") & "','�����¼','�޸ĳ�ֵ��ID:" & id & "')"
	rs.close
	call smsg("��ϲ,�˳�ֵ�������޸����!","czclist.asp")
case "modinpc"
	call zzz()
	id=clng(request("id"))
	npcname=sqlstr(request.form("npcname"))
	npcsex=sqlstr(request.form("npcsex"))
	npcvalue=sqlstr(request.form("npcvalue"))
	npcimg=sqlstr(request.form("npcimg"))
	npcmoney=sqlstr(request.form("npcmoney"))
	npcgj=sqlstr(request.form("npcgj"))
	npcfy=sqlstr(request.form("npcfy"))
	npcwg=sqlstr(request.form("npcwg"))
	npctl=sqlstr(request.form("npctl"))
	npcnl=sqlstr(request.form("npcnl"))
	npcdie=sqlstr(request.form("npcdie"))
	npcgjl=sqlstr(request.form("npcgjl"))
	npccxl=sqlstr(request.form("npccxl"))
	npczhuren=sqlstr(request.form("npczhuren"))
	npcdiren=sqlstr(request.form("npcdiren"))
	npcwp=sqlstr(request.form("npcwp"))
	npcwplx=sqlstr(request.form("npcwplx"))
	npcccc=sqlstr(request.form("npcccc"))
	'��������!
	rs.Open "SELECT * FROM npc where id=" & id,conn,1,3
        if rs.eof or rs.bof then
		rs.close
		call smsg("�û�������!","-1")
	end if
	rs("n����")=npcname
	rs("n�Ա�")=npcsex
	rs("n����")=npcvalue
	rs("nͼƬ")=npcimg
	rs("n����")=npcmoney
	rs("n����")=npcgj
	rs("n����")=npcfy
	rs("n�书")=npcwg
	rs("n����")=npctl
	rs("n����")=npcnl
	rs("n������")=npcdie
	rs("n������")=npcgjl
	rs("n������")=npccxl
	rs("n����")=npczhuren
	rs("n����")=npcdiren
	rs("n��Ʒ")=npcwp
	rs("n��Ʒ����")=npcwplx
	rs("n������")=npcccc
	rs.Update
	conn.execute "insert into l(a,b,c,d,e) values (now(),'" & aqjh_name & "','" & Request.ServerVariables("REMOTE_ADDR") & "','�����¼','�޸�NPC���[" & id & "]')"
	rs.close
	call smsg("��ϲ,��NPC�����޸����!","npclist.asp")
End Select
sub zzz
if aqjh_name<>Application("aqjh_user") then 
	call smsg("�㲻����վ�����㲻�ܲ���!","-1")
end if
end sub
function sqlstr(mystr)
mystr=lcase(trim(mystr))
mystr=Replace(mystr,"union","u_nion")
mystr=Replace(mystr,"'","��")
mystr=Replace(mystr,"not","n_o_t")
if mystr="" then
	call smsg("���������Ϊ�ղ��ܲ���!","-1")
end if
sqlstr=mystr
end function
Sub smsg(str1, str2)
 set rs=nothing
 conn.close
 set conn=nothing
 onend = 0
 Select Case str2
 Case "-1","-2","-3" '����
    str2 = "location.href = 'javascript:history.go(-1)';"
 Case "1" '�ر�
    str2 = "window.close();"
 Case "2" '������
    str2 = ""
 Case "3" '����
    str2 = ""
    onend = 1
 Case Else '���˵�ָ���ļ�
    str2 = "location.href = '" & str2 & "';"
 End Select
 Response.Write "<script language=JavaScript>{alert('��ʾ��" & str1 & "');" & str2 & "}</script>"
 If onend = 0 Then
    Response.End
 End If
End Sub
%>
