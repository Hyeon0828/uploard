<!DOCTYPE html>
<html>
  <head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>방명록</title>
  </head>
  <style type="text/css">
    #div_tyme{
      display: flex;
      margin: 10px;
    }
    .box_txt1{
      font-weight: bolder;
      text-align: center;
      border: 2px ##000000 solid;
      width: 120px;
      font-size: 17px;
      border-right: 0px;
      border-radius: 10px 0px 0px 10px;
      background: #000000;
      color: white;
    }
    .box_txt2{
      border: #000000 solid;
      font-size: 16px;
      border-left: 0px;
      border-right: 0px;
      width: 400px;
    }
    .box_btn{
      border: 2px #000000 solid;
      font-size: 16px;
      background-color: #000000;
      color: white;
      font-weight: bolder;
      border-radius: 0px 10px 10px 0px;
      width: 80px;
    }
    .box_main{
      border: 3px #000000 solid;
      border-right: 0px;
      border-left: 0px;
      background: #2d2d2d;
    }
    .box_font{
      font-size: 17px;
      margin-right: 10px;
    }
  </style>
  <script language=javascript>
    function ok() {
      if (frma1.text_main.value.length == 0) {
        alert("내용을 적어주세요.")
        frma1.text_main.focus();
      }
      else {
        frma1.submit();
      }
    }
    function del() {
      window.open('guest_del.asp', '_self');
    }
  </script>
  <%
    Set db = Server.createObject ("ADODB.Connection")
    dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("bang.accdb")
    db.Open(dbMain)

    all_data_num = "select * from db_num"         '모든 방명록 개수 검색
    Set dummy_db = db.execute(all_data_num)

    db_count = int(dummy_db("db_datanum"))

    if db_count mod 6 = 0 then all_page_num = db_count\6        '페이지 쪽수
    if db_count mod 6 <> 0 then all_page_num = db_count\6 + 1   '게시물을 6개넘어가면 쪽수 추가
  %>
  <body onload="frma1.text_main.focus();">
   <form name="frma1" method="post" action="guest_insert.asp">
    <center>
    <table>
      <tr><td>
        <div id="div_tyme">
          <input type="text" name="text_name" size="6" maxlength="6"class="box_txt1" readonly
           value=<% if session("ss_name") = "" then
                      response.write "비회원"
                    else
                      response.write session("ss_name")
                    end if %>
           >
          <textarea name="text_main" rows="3" class="box_txt2" maxlength="70" placeholder="비방, 욕설 등을 작성하는 것을 목격하면 ip차단합니다."></textarea>
          <input type="button" value="작성하기" onClick="ok();" class="box_btn" name="btn_1"
          onmouseover="on_btn();" onmouseout="out_btn();">
        </div>
        </td></tr>
     </table>
    </form>
    <iframe src="guest_list.asp?page=<%=all_page_num%>" width="700" height="410" class="box_main" name="iframe_list"></iframe><br>
    <%
      response.write "<font class=box_font><</font>"
      if db_count mod 6 = 0 then
        for i = 1 to db_count/6
          response.write "<a href='guest_list.asp?page=" & all_page_num - i + 1 & "' target='iframe_list' class=box_font>" & i & "</a>"
        next
      else
        for i = 1 to db_count/6 + 1
          response.write "<a href='guest_list.asp?page=" & all_page_num - i + 1 & "' target='iframe_list' class=box_font>" & i & "</a>"
        next
      end if
      response.write "<font class=box_font style='margin:0px;margin-left:2px;'>></font>"

      db.close()
      set db = nothing
    %>
  </center>
    <br><% if session("ss_m") = "m" then
    response.write "<input type=button value='Reset' class='btn btn-xs btn-danger'" &_
    "onClick='del();'>"
    end if %>
  </body>
</html>
