<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
</head>

<body>
<%dim a,b,i,j,sex,xingming,strsql,address
strsql="select * from shuju where =sex"
xingming=request.Form("xingming")
address=request.Form("address")
if xingming<>"" then
strsql=strsql&"and xingming like '%"&xingming&"%'"
end if
if address<>"" then
strsql=strsql&" and address like '%"&address&"%'"

end if

set a=server.CreateObject("adodb.connection")
a.open"driver=sql server;server=pc07;uid=sa;pwd=;database=shuju"
set b=server.CreateObject("adodb.recordset")
b.open strsql,a,3,3
j=b.recordcount
%>





<form action="" method="post"><table width="921" border="1" align="center">
  <tr>
    <td colspan="2">&nbsp;</td>
    </tr>
  <tr>
    <td>��ַ</td>
    <td><label>
      <input name="address" type="text" id="address" value="" />
    </label></td>
  </tr>
  <tr>
    <td width="456">����</td>
    <td width="449"><label>
      <input name="xingming" type="text" id="xingming" value="" />
    </label></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><label>
     
     
        <input type="submit" name="Submit" value="�ύ" />
        &nbsp;&nbsp;
    </label>
	 
	    <input type="reset" name="Submit" value="����" />
	    </label>     </td>
    </tr>
</table>
</form>
<table width="921" height="112" border="1" align="center" >
  <tr>
    <td>ѧ��</td>
    <td height="58">����</td>
    <td>�Ա�</td>
    <td>��ַ</td>
    <td>����</td>
    <td>�绰</td>
    <td>����</td>
  </tr>
 <% for i=1 to j
 if b("sex") then
 	sex="��"
else
	sex="Ů"
end if
 
 
 %>
  <tr>
    <td><%=b("code")%></td>
    <td height="46"><%=b("xingming")%></td>
    <td><%=sex%>	</td>
    <td><%=b("address")%></td>
    <td><%=b("intro")%></td>
    <td><%=b("tel")%></td>
    <td><%=b("age")%></td>
  </tr>
  <%b.movenext
  next
  
  %>
</table>
<p>&nbsp;</p>
<p>strsql=strsql&amp;&quot;and name='&quot;&amp;xingming&amp;&quot;'&quot;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>b.open&quot;select*from shuju&quot;,a,3,3</p>
<p>&nbsp;</p>
</body>
</html>
