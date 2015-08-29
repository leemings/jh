<window name="tq_w";showinfo=10000;debug=0;speed=10>

<class:tools.jh.jh:jh delay=300;showsocket2=1;/>
<jh opendb="com.ms.jdbc.odbc.JdbcOdbcDriver,'admin','',jdbc:odbc,MData,user"/>

<class:xxtgames.xggroup:varchatgroup />

<class:games.bout.billiards.CBilliardsDesk:CBilliardsDesk9;
	sockets="b9_s";
	type=9;
	new=1;
/>
<class:games.bout.billiards.CBilliardsUserJh:CBilliardsUser9;
	groupclass=CBilliardsGame9;
	ping=1;
	delay=300;
	disableselfthread=1;
/>

<class:games.bout.billiards.CBilliardsDesk:CBilliardsDesk16;
	sockets="b16_s";
	type=16;
	new=1;
/>
<class:games.bout.billiards.CBilliardsUserJh:CBilliardsUser16;
	groupclass=CBilliardsGame16;
	ping=1;
	delay=300;
	disableselfthread=1;
/>

<sockets APPNAME="b9_s";name="b9_s";userclass=CBilliardsUser9;port=10004;/>

<sockets APPNAME="b16_s";name="b16_s";userclass=CBilliardsUser16;port=10005;/>

#define TBUSERDB  "MData,user,dataex,lv,sc,mt,wn"

<gsockets APPNAME="tq_rs";userclass=jh;port=10003;share="#(GMG,%name,水晶晶台球,0)";>

<class:games.bout.billiards.CBilliardsGameJh:CBilliardsGame9;
	manclass="games.bout.billiards.CBilliardsMan";
	mode="9";
	usermax=16;
	playmax=2;
	master=1;
	new=1;
	moneymatch=0;
	permissions=0;
/>
<varchatgroup userdbdef=#[TBUSERDB];clinedef="e%parent,in,%name,九球游戏室(免费/局),00">
<file include="billiards9_00.fml"/>
</varchatgroup>
<class:games.bout.billiards.CBilliardsGameJh:CBilliardsGame9;
	manclass="games.bout.billiards.CBilliardsMan";
	mode="9";
	usermax=16;
	playmax=2;
	master=1;
	new=1;
	moneymatch=5;
	permissions=0;
/>
<varchatgroup userdbdef=#[TBUSERDB];clinedef="e%parent,in,%name,九球游戏室(5元/局),00">
<file include="billiards9_00.fml"/>
</varchatgroup>

<class:games.bout.billiards.CBilliardsGameJh:CBilliardsGame16;
	manclass="games.bout.billiards.CBilliardsMan";
	mode="16";
	usermax=16;
	playmax=2;
	master=1;
	new=1;
	moneymatch=0;
	permissions=0;
/>
<varchatgroup userdbdef=#[TBUSERDB];clinedef="e%parent,in,%name,十六球游戏室(免费/局),00">
<file include="billiards16_00.fml"/>
</varchatgroup>
<class:games.bout.billiards.CBilliardsGameJh:CBilliardsGame16;
	manclass="games.bout.billiards.CBilliardsMan";
	mode="16";
	usermax=16;
	playmax=2;
	master=1;
	new=1;
	moneymatch=5;
	permissions=0;
/>
<varchatgroup userdbdef=#[TBUSERDB];clinedef="e%parent,in,%name,十六球游戏室(5元/局),00">
<file include="billiards16_00.fml"/>
</varchatgroup>
</gsockets>
</window>
