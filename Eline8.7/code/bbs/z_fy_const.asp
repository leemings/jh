<%
dim content,towho,mysign
dim memo_set,bs_set,tj_set,pst_set,fj_set,lh_set,qj_set,fymaster,jymaster,log_set
dim fjty_set,fjbh_set,fjpf_set,fjzy_set,fjwg_set,fyname,jyname
sub logs(lclass,title,userN,tousername)
       if UserID="" or (not isnumeric(UserID)) then UserID=0
	if content="" then content="©ири╣д╡ывВ"
	connfy.execute("insert into log ([class],UserName,UserID,Title,content,DateAndTime,Tousername,sign,IP) values('"&lclass&"','"&userN&"',"&UserID&",'"&Title&"','"&content&"',now(),'"&tousername&"','"&mysign&"','"&Request.ServerVariables("REMOTE_ADDR")&"')")
end sub
sub bnew(title,content,username)
          conn.execute("insert into [bbsnews] (boardid,title,content,username,addtime) values (0,'"&title&"','"&content&"','"&username&"',now())")
end sub
sub messto(towho,sender,title,content)
    conn.execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&towho&"','"&sender&"','"&title&"','"&content&"',Now(),0,1)")
end sub
sub getfyconfig()
	set rs=connfy.execute("select * from fyconfig")
	memo_set=checkstr(rs(0))
	bs_set=split(rs(1),"|")
	tj_set=split(rs(2),"|")
	pst_set=split(rs(3),"|")
	fj_set=split(rs(4),",")
	lh_set=split(rs(5),"|")
	qj_set=split(rs(6),"|")
	fyname=checkstr(trim(rs(7)))
	jyname=checkstr(trim(rs(8)))
	log_set=split(rs(9),"|")
	fjty_set=split(fj_set(0),"|")
	fjbh_set=split(fj_set(1),"|")
	fjpf_set=split(fj_set(2),"|")
	fjzy_set=split(fj_set(3),"|")
	fjwg_set=split(fj_set(4),"|")
	if membername=fyname then fymaster=true
	if membername=jyname then jymaster=true
end sub
%>