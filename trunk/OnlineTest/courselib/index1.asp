<!--#include file="connection.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />

<title>�γ����</title>

<link rel="stylesheet" href="../link.css">
<style type="text/css">
<!--
.STYLE8 {color: #003300; font-size: 24pt; font-weight: bold; font-family: "����"; }
.STYLE10 {color: #006600; font-size: 14; font-family: "����"; }
	.word{
	size:4px;
	font-family:"����";}
	.top{
	vertical-align:top;}
	.btn{
	background-color:#CCFFCC;
	width:130px;
	height:30px;
	}


-->
</style>
<script src="../../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<body topmargin="0" leftmargin="0" marginheight="0" marginwidth="0" background="image/page_bg.jpg">

		
		<table border="0" cellpadding="0" cellspacing="0" width="950">
		<tr>
			<td><a href="../../index.htm"><img src="image/logo.gif" border="0"></a></td>
			<td><script type="text/javascript">
AC_FL_RunContent( 'codebase','http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0','width','400','height','90','id','index','align','middle','src','../../swf/index?pageNum=0','quality','high','bgcolor','#ffffff','name','index','allowscriptaccess','sameDomain','pluginspage','http://www.macromedia.com/go/getflashplayer','movie','../../swf/index?pageNum=0','wmode','transparent' ); //end AC code
</script><noscript><object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" name="index" width="400" height="90" align="middle" id="index">
<param name="allowScriptAccess" value="sameDomain" />
<param name="movie" value="../../swf/index.swf?pageNum=0" />
<param name="quality" value="high" />
<param name="bgcolor" value="#ffffff" />
<param name="wmode" value="transparent">
<embed src="../swf/index.swf?pageNum=0" quality="high" bgcolor="#ffffff" width="400" height="90" name="index" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />
</object></noscript></td>
			<td><script type="text/javascript">
AC_FL_RunContent( 'codebase','http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0','width','400','height','90','id','index','align','middle','src','../../swf/index1?pageNum=0','quality','high','bgcolor','#ffffff','name','index','allowscriptaccess','sameDomain','pluginspage','http://www.macromedia.com/go/getflashplayer','movie','../../swf/index1?pageNum=0','wmode','transparent' ); //end AC code
</script><noscript><object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" name="index" width="400" height="90" align="middle" id="index">
<param name="allowScriptAccess" value="sameDomain" />
<param name="movie" value="../../swf/index1.swf?pageNum=0" />
<param name="quality" value="high" />
<param name="bgcolor" value="#ffffff" />
<param name="wmode" value="transparent">
<embed src="../swf/index1.swf?pageNum=0" quality="high" bgcolor="#ffffff" width="400" height="90" name="index" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />
</object></noscript></td>
            <td><img src="image/top_menu.gif" border="0" usemap="#topmenumap"></td>
		</tr>

		</table>
		
		<map name="topmenumap">
		  <area shape="rect" coords="31,19,111,28" href="../../index.htm" target="_self">
		  <area shape="rect" coords="35,46,113,56" href="#">
		</map>
	

	
		<table border="0" cellpadding="0" cellspacing="0" width="950"  >
		<tr>
			<td width="219" valign="top"  height="100%">
			
				<table border="0" cellpadding="0" cellspacing="0"  height="100%">
				<tr>
					<td><img src="image/menu_top.gif"></td></tr>
				<tr>
					<td><a href="../qualif/index.htm"><img src="image/left_menu01.gif" border="0" ></a></td>
				</tr>
				<tr>
					<td><a href="index.htm"><img src="image/left_menu02.gif" border="0"  ></a></td>
				</tr>
				<tr>
					<td><img src="image/left_menu03.gif" border="0"  ></td>
				</tr>
				<tr>
					<td><img src="image/left_menu04.gif" border="0"  ></td>
				</tr>
				<tr>
					<td><a href="../qualif/index.htm"><img src="image/left_menu05.gif" border="0" ></a></td>
				</tr>
				<tr>
					<td><img src="image/left_menu06.gif" border="0"  ></td></tr>
				<tr>
					<td><img src="image/left_menu_bg.gif"></td></tr>
				<!--/??-->
				<!--??? ???-->
				
				<tr>
					<td><img src="image/page_img.jpg"></td></tr>
				<tr>
					<td height="100%" background="image/page_menu_bg.gif"></td></tr>
				</table>
				
			</td>
			

			
			<td width="731" valign="top" colspan="2">
				
				<table border="0" cellpadding="0" cellspacing="0" width="731">
				<tr>
					<td><img src="image/page_top.gif"></td></tr>
				<tr>
					<td bgcolor="#ffffff" align="center"><img src="image/page_title.gif"></td></tr>
				</table>
			
				
				
				<table  border="0" cellpadding="0" cellspacing="0" width="731">
				<tr>
				
					<td width="731" align="center" valign="top">
					<br><%
			dim sql,i,num
	sql="select * from chooseTwo"
	rs.open sql,conn,1,1
	num=1
	
	'��ʾ��⣺
	response.Write "<form action='' method='post'>"
	response.write "<table border=0 cellpadding='3' cellspacing='3' align='center'>"
	Do While Not rs.eof
	    response.write "<tr>"
		response.write "<td colspan='4' bgcolor='#CCFFCC'><font class='word'><b>"&num&"."&rs(1)&"</b></font></td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td><input type='radio' name='same"&num&"' value='A'/><font class='word'>"&"A."&rs(2)&"</font></td>"
		response.write "<td><input type='radio' name='same"&num&"' value='B'/><font class='word'>"&"B."&rs(3)&"</font></td>"
		response.write "<td><input type='radio' name='same"&num&"' value='C'/><font class='word'>"&"C."&rs(4)&"</font></td>"
		response.write "<td><input type='radio' name='same"&num&"' value='D'/><font class='word'>"&"D."&rs(5)&"</font></td>"
		response.write "</tr>"
		num=num+1
	    rs.movenext
	Loop 
	response.Write "<tr><td colspan='2' align='center'><input type='submit' value='����' class='btn'></td>"
	response.Write "<td colspan='2' align='center'><input type='reset' value='����' class='btn'></td></tr>"
	response.write "</table>"
	response.Write "</form>"
	num=num-1
%>
<%

			function check()
				for i=1 to num
					if request("same"&i)="" then
						check=0
					else
						check=1
						exit for
					end if
				next
			end function

	'��������a()��ѧ������
	'        b()����ȷ��
	if check()=1 then
	set rs=conn.execute(sql)
	dim a(20),b(20),j
	j=1
	response.Write "���Ĵ������ǣ�<br>"
	for i=1 to num
		a(i)=request.Form("same"&i)
		response.Write i&"."&a(i)&"&nbsp;&nbsp;&nbsp;&nbsp;"
	next
	response.Write "<br>====================================================================<br>"
	Do While Not rs.eof
		b(j)=rs(6)
		j=j+1
		rs.movenext
	Loop
	
	'��ʾ���
	j=1
	response.Write "<br>���������Ŀ��:"
	for i=1 to num
		if a(i)<>b(j) then
			response.Write i&";"
		end if
		j=j+1
	next
	response.Write "<br>"
	
	response.Write "<br>��ȷ�𰸣�<Br>"
	for j=1 to num
			response.Write j&"."&b(j)&"&nbsp;&nbsp;&nbsp;"
	next
	End if
%></td>
					
					
				
				</tr>
				</table>
				
			</td></tr>
			



		</table>
		

<!--Copyright-->
<table border="0" cellpadding="0" cellspacing="0" width="100%" bgcolor="#F7F4E7">
<tr>
	<td width="20"></td>
	<td align="left"><img src="image/end_box.gif"></td></tr>
</table>
<!--/Copyright-->
</body>
</html>
<span style="display:none;"><a href="http://www.mobanwang.com" title="ؕŮ??">ؕŮ??</a></span>