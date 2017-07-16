<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>

<body>
<%dim a,b,i,j,sex,xingming,strsql,address，code
strsql="select * from shuju where code=code"


set a=server.CreateObject("adodb.connection")
a.open"driver=sql server;server=pc59;uid=sa;pwd=;database=shuju"
set b=server.CreateObject("adodb.recordset")
b.open "select*from shuju",a,3,3
j=b.recordcount

xingming=request.Form("xingming")
address=request.Form("address")
if xingming<>"" then
strsql=strsql&"and xingming like '%"&xingming&"%'"
end if
if address<>"" then
strsql=strsql&" and address like '%"&address&"%'"

end if
%>
<form action="" method="post">
  <table width="921" border="1" align="center">
    <tr>
      <td colspan="2"><div align="center">查询条件</div></td>
    </tr>
    <tr>
      <td>地址</td>
      <td><label>
        <input name="address" type="text" id="address" />
      </label></td>
    </tr>
    <tr>
      <td width="456">姓名</td>
      <td width="449"><label>
        <input name="xingming" type="text" id="xingming" value="" />
      </label></td>
    </tr>
    <tr>
      <td colspan="2" align="center"><label>
        <input type="submit" name="Submit" value="提交" />
        &nbsp;&nbsp; </label>
          <input type="reset" name="Submit" value="重置" />
          </label>
      </td>
    </tr>
  </table>
</form>
<table width="830" height="196" border="1" align="center" >
  <tr>
    <td width="85">学号</td>
    <td width="85" height="93">姓名</td>
    <td width="85">性别</td>
    <td width="85">地址</td>
    <td width="85">电话</td>
    
    <td width="85">年龄</td>
	<td width="92">&nbsp;</td>
  </tr>
  <% for i=1 to j
 if b("sex") then
 	sex="男"
else
	sex="女"
end if
 
 
 %>
  <tr>
    <td><%=b("code")%></a></td>
    <td height="95"><a href="file:///E|/新建文件夹/eb1/41 lifu/asp/shujuku3.asp?code=<%=b("code")%>"><%=b("xingming")%></a></td>
    <td><%=sex%> </td>
    <td><%=b("address")%></td>
    <td><%=b("tel")%></td>
    
    <td><%=b("age")%></td>
	<td>修改 <a href="file:///E|/新建文件夹/eb1/41 lifu/asp/shujukuzhuc3.asp?code=<%=b("code")%>">删除</a></td>
  </tr>
  <%b.movenext
  next
  
  %>
</table>

<p>&nbsp;</p>
</body>
</html>
