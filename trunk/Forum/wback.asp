<!--#include file="conn.asp"--> 
<link href="main.css" rel="stylesheet" type="text/css" />
<%
idh=Replace(Request.Form("id_h"),"'","''")
nameb=Replace(Request.Form("name_b"),"'","''") 
titleb=Replace(Request.Form("title_b"),"'","''") 
bodyb=Replace(Request.Form("body_b"),"'","''") 
codeb=Replace(Request.Form("code_b"),"'","''")
set rs=conn.execute("select * from author where name='" &nameb& "' and code='" &codeb& "'")
sql="select * from author where name='" &nameb& "' and code='" &codeb& "'"
if nameb="" or titleb="" or bodyb="" or idh="" or rs.eof then%><title>�ظ�����</title>
������Ϊ��/������û���������!<a href="show.asp?id=<%=idh%>">�鿴����</a>,���<a href="register.asp">ע��</a>
<%else
sql="update bbs set countwb=countwb+1 where id=" &idh
conn.execute(sql)
sql="insert into bbbs(hostid,btime,bname,bcontent,btitle) " & _
         "values('"& idh & "','" & now() &"','"& nameb &"','"& bodyb &"','"& titleb &"')"
conn.execute(sql)%> 
����ɹ���<a href="index.asp">�ص���ҳ</a> <br>
<a href="show.asp?id=<%=idh%>">�鿴����</a>
<%response.Redirect("show.asp?id="&idh)
end if %>
