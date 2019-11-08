<HTML>
<head>
</head>
 <script id=clientEventHandlersJS language=javascript>
  function Button1_onclick()
  {
   if(frm.adr.value == "")
   {
    alert("주소를 입력해주세요.");
    return;
   }
  frm.submit();
  }

function address_ok()
  {
   var text_main = frm_zip.sel_address.value;
   var text_1 = text_main.substr(0, 3);
   var text_2 = text_main.substr(4, 3);
   var text_3 = text_main.substr(8, 30);

   opener.main.zip1.value = text_1;
   opener.main.zip2.value = text_2;
   opener.main.zip3.value = text_3;

   self.close();
   }

   function cancle() {
   self.close();
   }

</script>

<body onload="frm_zip.text_search.focus()">
<center>
 <div id="div_back">
  <div id="div_form">
  <font size=4><b>주소 검색</b><p></font>
  <form method="post" name="frm_zip" action="info_zip.asp">
  <table>
   <tr>
   <td> <input type="text" name="text_search" maxlength="20"></td>
   <td> <input type="submit" value="검색" name="name_sc"></td>
   </tr>
   </table><p>

<%
dim address, search_address, sql
address = request("text_search")

 if(address = "") then
  response.write "<b>검색할 주소(동)를 입력하세요.</b>" &_
  "<input type='button' value='취소' onclick='cancle();'>"

  else
  Set db = Server.createObject ("ADODB.Connection")
    dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("zipcode.mdb")
    db.Open (dbMain)

  sql = "select * from zipcode where GUGUN = '" & address & "'" & " or DONG like '%" & address & "%' "
  set search_address = db.execute(sql)

  if search_address.EOF then
   response.write"<font color=black><b>" & id & "</b></font><p><b><font>찾을 수 없는 </b></font>주소입니다.<p>" &_
   "<input type='button' value='취소' onclick='cancle();'>"

   else
    response.write "<font><b>찾는 주소를 선택하세요. </b></font><p>"
    response.write "<select name='sel_address'>"
      do while not search_address.EOF
      dim index
      index = search_address("ZIPCODE") & " " & search_address("SIDO") & " " & search_address("GUGUN") & " " & search_address("DONG") & " " & search_address("BUNJI")
      response.write "<option value='" & index & "'>" & index & "</option>"
      search_address.MoveNext
      Loop
      response.Write "</select><p><input type='button' value='취소' OnClick='cancle();'> &nbsp <input type='button' value='선택' OnClick='address_ok();'>"
        end if
       db.close()
     set db = nothing
     set search_address = nothing
     end if
%>
   </form>
 </center>
 </div>
 </div>
  </body>
 </html>
