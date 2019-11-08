<html>
<head>
<title>회원정보검색</title>
</head>

<%
ss_id = Session("user_id")
    Set db = Server.createObject ("ADODB.Connection")
    con = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("ex.accdb")
    db.Open (con)

    Dim one_db, sql
    sql = "select * from ex1table where db_id ='" & ss_id & "'"
    Set one_db = db.execute(sql)

    ss_gen = one_db("db_gender")

 ss_zip1 = one_db("db_zip1")
 ss_zip2 = one_db("db_zip2")
 ss_zip3 = one_db("db_zip3")
 ss_zip4 = one_db("db_zip4")
 ss_call1 = one_db("db_phone1")
 ss_call2 = one_db("db_phone2")
 ss_call3 = one_db("db_phone3")
 ss_mail1 = one_db("db_email1")
 ss_mail2 = one_db("db_email2")


%>

<script language=javascript>
function submit_1() {
if(main.phone1.value<3) {
alert("전화번호를 입력해주세요.");
document.main.phone1.focus();
return;
}

else if(main.phone2.value<4) {
  alert("전화번호를 입력해주세요.");
  document.main.phone2.focus();
  return;
}

else if(main.phone3.value<4) {
  alert("전화번호를 입력해주세요.");
  document.main.phone3.focus();
  return;
}

else {
  main.submit();
}
}

function find() {
  window.open('info_zip.asp', '', 'height=450px, width=400px');
}

function nospace() {
  if(event.keyCode == 32) {
    event.keycode = 0;
    return false;
  }
}
document.onkeydown = nospace;

function fncInputKorean() {
    if ((event.keyCode < 12592) || (event.keyCode > 12687)) {
        event.keycode = 0;
        event.returnValue = false;

    }
}


function submit_2() {
 window.open('info_del.html', '', 'status=no, width=650px, height=400px')
}

function cancle() {
  self.close();
}




</script>

<style>
top {
  width : 100%
  height : 50px;
  background-color: #2d2d2d;
  padding:0px;
}
</style>


<body>
  <form name="main" method="post" action="info_updat.asp">
    <div class="top">
      <center>
        <fornt color="black" size="6">회원정보</font>
      </center>
    </div>

    <table>
      <tr>
        <td>User ID : </td>
        <td><% response.write one_db("db_id") %></td>
      </td>

      <tr>
        <td>Name</td>
        <td><input type="text" size="10" name="text_name" onkeypress="fncInputKorean()" maxlength="6" value= "<% response.write one_db("db_name") %>"></td>
      </tr>

      <tr>
        <td>Gender</td>
        <td><% response.write one_db("db_gender") %></td>
      </tr>

      <tr>
        <td>Phone</td>
        <td><input type="text" size="4" name="phone1" onKeyup="this.value=this.value.replace(/[^0-9]/g,'');" maxlength="3" value="<% response.Write one_db("db_phone1") %>"> - <input type="text" size="4" maxlength="4" name="phone2" onKeyup="this.value=this.value.replace(/[^0-9]/g,'');" value="<% response.Write one_db("db_phone2") %>"> -
          <input type="text" size="4" maxlength="4" name="phone3" value="<% response.Write one_db("db_phone3") %>" onKeyup="this.value=this.value.replace(/[^0-9]/g,'');">
        </td>
      </tr>

      <tr>
        <td>e-mail</td>
        <td><input type="text" size="10" maxlength="10" name="email" onkeypress="nospace();" value= "<% response.write one_db("db_email1") %>">@
        <input type="text" size="15" maxlength="20" name="domain" onkeypress="nospace();" value="<% response.write one_db("db_email2") %>">
        </td>
      </tr>

      <tr>
        <td>주소</td>
        <td><input type="text" size="5" readonly name="zip1" readonly value="<% response.write one_db("db_zip1") %>">
            <input type="text" size="5" readonly name="zip2" readonly value="<% response.Write one_db("db_zip2") %>">
            <input type="button" value="검색" name="find_zip" onclick="find();">
        </td>
      </tr>

      <tr>
        <td>
        </td>
        <td><input type="text" name="zip3" readonly size="40" value="<% response.write one_db("db_zip3") %>"></td><br>
      </tr>
      <tr>
        <td></td>
        <td><input type="text" name="zip4" size="40" value="<% response.write one_db("db_zip4") %>"></td><br>
      </tr>

      <tr>
        <td><input type="button" value="취소" onclick="cancle();"></td>
        <td align="right">
          <input type="button" value="수정" onclick="submit_1();"> <input type="button" value="삭제" onclick="submit_2();">
        </td>
      </tr>
</table>
</form>
</body>


<% db.close() %>
</html>
