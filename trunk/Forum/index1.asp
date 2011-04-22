<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="header.htm"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />

<title>讨论区</title>

<link rel="stylesheet" href="../link.css">
<style type="text/css">
<!--
.STYLE8 {color: #003300; font-size: 24pt; font-weight: bold; font-family: "黑体"; }
.STYLE10 {color: #006600; font-size: 14; font-family: "黑体"; }
	.top{
	vertical-align:top;}
	.btn{
	background-color:#CCFFCC;
	width:130px;
	height:30px;
	}


-->
</style>
<script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<body topmargin="0" leftmargin="0" marginheight="0" marginwidth="0" background="image/page_bg.jpg">
<div title="123">



		
		<table border="0" cellpadding="0" cellspacing="0" width="950">
		<tr>
			<td><a href="../../index.htm"><img src="image/logo.gif" border="0"></a></td>
			<td>
			<script type="text/javascript">
AC_FL_RunContent( 'codebase','http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0','width','400','height','90','id','index','align','middle','src','../swf/indexs?pageNum=0','quality','high','bgcolor','#ffffff','name','index','allowscriptaccess','sameDomain','pluginspage','http://www.macromedia.com/go/getflashplayer','movie','../swf/indexs?pageNum=0','wmode','transparent' ); //end AC code
</script><noscript><object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" name="index" width="400" height="90" align="middle" id="index">
<param name="allowScriptAccess" value="sameDomain" />
<param name="movie" value="../swf/indexs.swf?pageNum=0" />
<param name="quality" value="high" />
<param name="bgcolor" value="#ffffff" />
<param name="wmode" value="transparent">
<embed src="../swf/indexs.swf?pageNum=0" quality="high" bgcolor="#ffffff" width="400" height="90" name="index" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />
</object></noscript></td>
            <td><script type="text/javascript">
AC_FL_RunContent( 'codebase','http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0','width','400','height','90','id','index','align','middle','src','../swf/index1s?pageNum=0','quality','high','bgcolor','#ffffff','name','index','allowscriptaccess','sameDomain','pluginspage','http://www.macromedia.com/go/getflashplayer','movie','../swf/index1s?pageNum=0','wmode','transparent' ); //end AC code
</script><noscript><object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" name="index" width="400" height="90" align="middle" id="index">
<param name="allowScriptAccess" value="sameDomain" />
<param name="movie" value="../swf/indexs1.swf?pageNum=0" />
<param name="quality" value="high" />
<param name="bgcolor" value="#ffffff" />
<param name="wmode" value="transparent">
<embed src="../swf/indexs1.swf?pageNum=0" quality="high" bgcolor="#ffffff" width="400" height="90" name="index" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />
</object></noscript></td>
			<td><img src="image/top_menu.gif" border="0" usemap="#topmenumap"></td>
		</tr>

		</table>
		
		<map name="topmenumap">
		  <area shape="rect" coords="31,19,111,28" href="../index.htm" target="_self">
		  <area shape="rect" coords="35,46,113,56" href="#">
		</map>
	

	
		<table border="0" cellpadding="0" cellspacing="0" width="950"  >
		<tr>
			<td width="219" valign="top"  height="100%">
			
				<table border="0" cellpadding="0" cellspacing="0"  height="100%">
				<tr>
					<td><img src="image/menu_top.gif"></td></tr>
				<tr>
					<td><img src="image/left_menu01.gif" border="0" ></td>
				</tr>
				<tr>
					<td><img src="image/left_menu02.gif" border="0"  ></td>
				</tr>
				<tr>
					<td><img src="image/left_menu03.gif" border="0"  ></td>
				</tr>
				<tr>
					<td><img src="image/left_menu04.gif" border="0"  ></td>
				</tr>
				<tr>
					<td><img src="image/left_menu05.gif" border="0" ></td>
				</tr>
				<tr>
					<td><img src="image/left_menu06.gif" border="0"  ></td></tr>
				<tr>
					<td><img src="image/left_menu_bg.gif"></td></tr>
				
				
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
					
					<!--<div align="center">-->
					
<%sub maintalk()%>
<%
set showbbs = server.CreateObject("ADODB.recordset")
showbbsStr="select*from bbs order by wtime desc"
showbbs.open showbbsStr,conn,3,2
if showbbs.EOF and showbbs.BOF then%>
<br><center>暂时还没有文章,现在<a href="say.asp">发表</a></center>
<%else
showbbs.PageSize=10'分页
PageN=10 '显示10个页数.
PageCount=showbbs.PageCount
Page=int(request("Page"))
CurrentPageN=int(request("CurrentPageN"))
if Page<=0 or request("Page")="" or request("Page")="0" then Page=1
if CurrentPageN<=0 or request("CurrentPageN")="" then CurrentPageN=1
showbbs.AbsolutePage=Page
%>
<br>
<%for i=1 to showbbs.PageSize%>
<!--<div id="arttitle">-->
      <a href="show.asp?id=<%=showbbs("id")%>"><%=showbbs("title")%></a><br><br>
<!--</div>-->

<!--<div id="authorn">-->
      作者:<a href="authordetail.asp?auname=<%=showbbs("name")%>"><%=showbbs("name")%></a>&nbsp;
      <%if showbbs("countwb")=0 then%>
       回复&lt;<%=showbbs("countwb")%>&gt;<!--回复条数 -->
      <%else%>
       <a href="show.asp?id=<%=showbbs("id")%>">
       回复&lt;<font color=#ff0000><b><%=showbbs("countwb")%></b></font>&gt;
       </a><!--回复条数 -->
      <%end if%>&nbsp;
  
      <%wt=showbbs("wtime")'如果为撰写当天的时间，红色显示。
      if year(now())=year(wt) and month(now())=month(wt) and day(now())=day(wt) then%>
       撰写/修改时间:<font color=#ff0000><%=wt%></font>&nbsp;
      <%else%>
       撰写/修改时间:<%=wt%>&nbsp;
      <%end if%>
  
      <a href="modify.asp?id=<%=showbbs("id")%>">修改</a>
      <a href="del.asp?id=<%=showbbs("id")%>">删除</a>
<!--</div>-->

<hr size="1" width=600  align="left" style=" border:dashed #999999 ">

<%showbbs.movenext
if showbbs.EOF then exit for
Next%>
<center>
<br>
<%
if Page<>1 then%>
<a href="index.asp?Page=1&CurrentPageN=1">(首页)</a>
<%else%>
(当前为首页)
<%end if%>
<%if PageCount>PageN and CurrentPageN>=2 then%>
<a href="index.asp?Page=<%=(CurrentPageN-2)*PageN+1%>&CurrentPageN=<%=CurrentPageN-1%>">(上<%=PageN%>页)</a>
<%end if%>
<%for i=(CurrentPageN-1)*PageN+1 to (CurrentPageN-1)*PageN+PageN%>
<%if i = Page then%>
      <font color=#ff0000 >-<%=i%>-</font>
<%elseif i<showbbs.RecordCount/showbbs.PageSize +1 then%>
      <a href="index.asp?page=<%=i%>&CurrentPageN=<%=CurrentPageN%>">&nbsp;<%=i%>&nbsp;</a>
<%end if%>
<%Next%>
<%if PageCount>PageN and CurrentPageN<PageCount/PageN then%>
<a href="index.asp?Page=<%=CurrentPageN*PageN+1%>&CurrentPageN=<%=CurrentPageN+1%>">(下<%=PageN%>页)</a>
<%end if%>
<%if PageCount>PageN and Page<>PageCount then%>
<a href="index.asp?Page=<%=PageCount%>&CurrentPageN=<%=int(PageCount/PageN)+1%>">(最后一页)</a>
<%end if%>
<%if PageCount=Page then%>(当前为最后页)<%end if%>
(共<%=PageCount%>页)
</center>
<%showbbs.Close
set showbbs=nothing
end if
end sub%>
<!------------------------------------------------- -->
<%sub saysth()
name=request.QueryString("name")%>
<hr size="1" width=700>
<form method="post" action="save.asp">
<table border="0" cellpadding="0" cellspacing="0" 
       style="border-collapse: collapse; " 
       width=600 align=center>
<tr><td>
发表文章:<br>
大名：<input type="text" name="name" size="23" value=<%=name%>>
密码：<input type="password" name="code" size="28" ID="Password1" style="font-size:12px">
点此<a href="register.asp">注册</a><br>
标题：<input type="text" name="title" size="67" ID="Text2"><br>
内容：<br>
<textarea rows="11" name="body" cols="72" ID="Textarea1"></textarea>
<br>
<input type="submit" value="提交" name="B1" ID="Submit1"><input type="reset" value="重置" name="B2" ID="Reset1">
<input type="hidden" name="name_b" size="20" ID="Hidden1"> <input type="hidden" name="title_b" size="28" ID="Hidden2">
</td></tr>
</table>
</form>
<%end sub%>


					<table width="770" border="0" cellspacing="0" align="center">
      <tr>
        <td id="tbul">　</td>
        <td id= "tbum"><%headtext()%></td>
        <td id="tbur">　</td>
      </tr>
      <tr>
        <td id="tbml">　</td>
        <td id="tbmm"><%maintalk()%><%saysth()%></td>
        <td id="tbmr">　</td>
      </tr>
  
      <tr>
        <td id="tbll">　</td>
        <td id="tblm">　</td>
        <td id="tblr">　</td>
      </tr>
</table>
					<!--</div>--></td>
					
					
				
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
</div>
</body>
</html>
<!--<span style="display:none;"><a href="http://www.mobanwang.com" title="貢女??">貢女??</a></span>-->