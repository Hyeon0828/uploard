<html>
<head>
</head>
<title>ID, PW ã��</title>
<script language=javascript>
 function cacle() {
  self.close();
 }

 function nospace() {
   if(event.keyCode == 32) {
     event.keycode = 0;
     return false;
   }
 }
 </script>

 <style>
 .button {

  width:48%;

  background-color: black;

  border: none;

  color:white;

  padding: 15px 0;

  text-align: center;

  text-decoration: none;

  display: inline-block;

  font-size: 15px;

  cursor: pointer;

padding-left:1px;
}
 </style>

<body onload="mainn.text_name.focus()">
 <div>
   <center>
     <a href="find_id.asp" class="button" target="_self">
      <font>ID�˻�</font></a>
     <a href="find_pw.asp" class="button" target="_self">
      <font>PW�˻�</font></a><br><br>
    <div>
     <form name="mainn" action="find_id.asp" method="post">
      <table>
      <tr>
      <td>����</td>
      <td><input type="text" name="text_name" maxlength="6" size="8" placeholder="�������" onkeypress="nospace()"><br></td>
      </tr>

      <tr>
      <td>�������</td>
      <td><input type="text" name="text_birth" maxlength="8" size="8" placeholder="ex)19000101" onkeypress="nospace()">       <input type="submit" class="button" value="�˻�"></td>
      </tr>
      </table>
    <p>
    </form>

<%
dim ss_name, ss_birth
ss_name = request("text_name")
ss_birth = request("text_birth")

if(ss_name = "") then
response.write "<p>" &_
             "</div><br><input type='button' value='���ư���' OnClick='cancle();'><br>"

 else
 Set db = Server.createObject ("ADODB.Connection")
  dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("ex.accdb")
  db.Open (dbMain)

dim search_id, sql
sql = "select * from ex1table where db_birth ='" & ss_birth & "'" & " and db_name ='" & ss_name & "'"
set search_id = db.execute(sql)

if search_id.EOF then
      Response.Write "<font color=#FF3636><b>ȸ�������� �����ϴ�.</b></font>" &_
      "<p>�ٽ� �Է����ּ���.<p>" &_
      "</div><br><input type='button' value='���ư���' OnClick='cancle();'><br>"

else
 response.write "<font><b>�˻� �� ID �Դϴ�.</b><font><p>"
  do while not search_id.EOF
   dim index
   index = search_id("db_id")
   response.write "<b><font size=4><li>" & index & "</li></font></b>"
  search_id.MoveNext
  Loop
  response.write "</div><br><input type='button' value='���ư���' onclick='cancle();'><br>"
end if
db.close()
set db = nothing
set search_id = nothing
end if
%>

</center>
</div>
</html>
