<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
</head>

<body><%dim code,strsql,sex
code=request.QueryString("code")
set a=server.CreateObject("adodb.connection")
a.open"driver=sql server;server=pc07;uid=sa;pwd=;database=shuju"
set b=server.CreateObject("adodb.recordset")
b.open "select * from shuju where code= '"&code&"'",a,3,3

sex=b("sex")
if sex="true" then
 	sex="��"
else
	sex="Ů"
end if

%>
<table width="962" border="1">
  <tr>
    <td colspan="2">������Ϣ</td>
  </tr>
  <tr>
    <td width="419">ѧ��</td>
    <td width="527"><%=b("code")%></td>
  </tr>
  <tr>
    <td>����</td>
    <td><%=b("xingming")%></td>
  </tr>
  <tr>
    <td>�Ա�</td>
    <td><%=sex)%></td>
  </tr>
  <tr>
    <td>��ַ</td>
    <td><%=b("address")%></td>
  </tr>
  <tr>
    <td>�绰</td>
    <td><%=b("tel")%></td>
  </tr>
  <tr>
    <td height="17">����</td>
    <td><%=b("age")%></td>
  </tr>
</table>

<p><a href="../index.html">������ҳ</a></p>
</body>
</html>
