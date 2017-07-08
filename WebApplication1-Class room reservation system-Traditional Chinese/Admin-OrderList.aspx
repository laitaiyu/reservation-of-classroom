<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
  <HEAD>
<!-- #include File="OleDbFunction.inc" -->
<meta content=True name=vs_snapToGrid>
<meta content="True" name=vs_showGrid>
<meta content="Microsoft Visual Studio .NET 7.1" name=GENERATOR>
<meta content="Visual Basic .NET 7.1" name=CODE_LANGUAGE>
<meta content=JavaScript name=vs_defaultClientScript>
<meta content=http://schemas.microsoft.com/intellisense/ie5 name=vs_targetSchema>
<Script Language="VB" Runat="Server">

  Dim holidays(12,31) as String

  Sub Load_Holidays()
      holidays(01,01) ="*元旦 中華民國開國紀念日"
      holidays(01,11) ="司法節"
      holidays(01,15) ="藥師節"
      holidays(01,23) ="自由日"
      holidays(02,04) ="農民節"
      holidays(02,14) ="情人節"
      holidays(02,15) ="戲劇節"
      holidays(02,28) ="*和平紀念日"
      holidays(03,01) ="兵役節"
      holidays(03,05) ="童子軍節"
      holidays(03,08) ="婦女節"
      holidays(03,12) ="植樹節"
      holidays(03,17) ="國醫節"
      holidays(03,20) ="郵政節"
      holidays(03,21) ="氣象節"
      holidays(03,25) ="美術節"
      holidays(03,26) ="廣播節"
      holidays(03,29) ="青年節"
      holidays(03,30) ="出版節"
      holidays(04,01) ="愚人節 主計節"
      holidays(04,04) ="婦幼節"
      holidays(04,05) ="音樂節"
      holidays(04,07) ="衛生節"
      holidays(04,22) ="世界地球日"
      holidays(05,01) ="*勞動節"
      holidays(05,04) ="文藝節"
      holidays(05,05) ="舞蹈節"
      holidays(05,10) ="珠算節"
      holidays(05,12) ="護士節"
      holidays(06,03) ="禁煙節"
      holidays(06,06) ="工程師節 水利節"
      holidays(06,09) ="鐵路節"
      holidays(06,15) ="警察節"
      holidays(06,30) ="會計師節"
      holidays(07,01) ="漁民節 公路節 稅務節"
      holidays(07,06) ="合作節"
      holidays(07,11) ="航海節"
      holidays(07,12) ="聾啞節"
      holidays(08,08) ="父親節"
      holidays(08,14) ="空軍節"
      holidays(09,01) ="記者節"
      holidays(09,03) ="軍人節"
      holidays(09,09) ="體育節 律師節"
      holidays(09,13) ="法律日"
      holidays(09,28) ="教師節"
      holidays(10,06) ="老人節"
      holidays(10,10) ="*國慶紀念日"
      holidays(10,21) ="華僑節"
      holidays(10,25) ="台灣光復節"
      holidays(10,31) ="萬聖節"
      holidays(11,01) ="商人節"
      holidays(11,11) ="工業節 地政節"
      holidays(11,17) ="自來水節"
      holidays(11,12) ="醫師節 中華文化復興節"
      holidays(11,21) ="防空節"
      holidays(12,05) ="海員節 盲人節"
      holidays(12,10) ="人權節"
      holidays(12,12) ="憲兵節"
      holidays(12,25) ="行憲紀念日 聖誕節"
      holidays(12,27) ="建築師節"
      holidays(12,28) ="電信節"
      holidays(12,31) ="受信節"  
  End Sub

  Sub Page_Load(sender As Object, e As Eventargs)
      If Session("LV") <> "Admin" Then
         Page.Response.Redirect ("./Main.aspx")
         Exit Sub
      End If
      Load_Holidays()
      If Not IsPostBack Then BindList()
  End Sub

  Sub BindList()
					 Dim strSQL As String 
 	  		strSQL = "Select [OrderMenu].[O-ID],[OrderMenu].[O-T_ID],[OrderMenu].[O-B_ID],[OrderMenu].[O-C_ID],[OrderMenu].[O-P_ID],[OrderMenu].[O-Date],[OrderMenu].[O-Time],[Teacher].[T-Name_LN_CHT],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name],[PhotoStudio].[P-Adds] From [OrderMenu],[Teacher],[Book],[Class],[PhotoStudio] " & _
				     	 		  "Where [OrderMenu].[O-Date] Like #" & DateString() & "#" & " And ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] And [OrderMenu].[O-B_ID] = [Book].[B-ID] And [OrderMenu].[O-C_ID] = [Class].[C-ID]  And [OrderMenu].[O-P_ID] = [PhotoStudio].[P-ID]) Order By [OrderMenu].[O-Date],[OrderMenu].[O-Time]"
					 myDataList.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "OrderMenu")
					 myDataList.DataBind()
  End Sub

  Sub BindList2()

				    Dim strSQL As String 
        If myCalendar.SelectedDates.Count = 1 Then
 	  			    strSQL = "Select [OrderMenu].[O-ID],[OrderMenu].[O-T_ID],[OrderMenu].[O-B_ID],[OrderMenu].[O-C_ID],[OrderMenu].[O-P_ID],[OrderMenu].[O-Date],[OrderMenu].[O-Time],[Teacher].[T-Name_LN_CHT],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name],[PhotoStudio].[P-Adds] From [OrderMenu],[Teacher],[Book],[Class],[PhotoStudio] " & _
				     	 		       "Where [OrderMenu].[O-Date] Like #" & DateValue(myCalendar.SelectedDate) & "#" & " And ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] And [OrderMenu].[O-B_ID] = [Book].[B-ID] And [OrderMenu].[O-C_ID] = [Class].[C-ID]  And [OrderMenu].[O-P_ID] = [PhotoStudio].[P-ID]) Order By [OrderMenu].[O-Date],[OrderMenu].[O-Time]"

        Else If myCalendar.SelectedDates.Count > 1 Then
           Dim strFirstDate As String 
           Dim strLastDate As String 
           With myCalendar.SelectedDates
                strFirstDate = .Item(0)
                strLastDate = .Item(.Count-1)
           End With
 	  			    strSQL = "Select [OrderMenu].[O-ID],[OrderMenu].[O-T_ID],[OrderMenu].[O-B_ID],[OrderMenu].[O-C_ID],[OrderMenu].[O-P_ID],[OrderMenu].[O-Date],[OrderMenu].[O-Time],[Teacher].[T-Name_LN_CHT],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name],[PhotoStudio].[P-Adds] From [OrderMenu],[Teacher],[Book],[Class],[PhotoStudio] " & _
				     	 		       "Where [OrderMenu].[O-Date] Between #" & strFirstDate & "# And #" & strLastDate & "#" & " And ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] And [OrderMenu].[O-B_ID] = [Book].[B-ID] And [OrderMenu].[O-C_ID] = [Class].[C-ID]  And [OrderMenu].[O-P_ID] = [PhotoStudio].[P-ID]) Order By [OrderMenu].[O-Date],[OrderMenu].[O-Time]"

        End If       
					   myDataList.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "OrderMenu")
					   myDataList.DataBind()

  End Sub

	 Sub DayChange( sender As Object, e As Eventargs )
				  Call BindList2()
	 End Sub
 
  Sub DataList_EditCommand(sender As Object, e As DataListCommandEventArgs)
      myDataList.EditItemIndex = e.Item.ItemIndex
      BindList()
  End Sub

  Sub DataList_CancelCommand(sender As Object, e As DataListCommandEventArgs)
      myDataList.EditItemIndex = -1
      BindList()
  End Sub

  Sub ExecuteSQL(strSQL As String)
      Dim objConn As New OleDbConnection()
      objConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                 "Data Source=" & Server.MapPath("./DB/order.mdb")
      objConn.Open()
      Dim objCmd As New OleDbCommand(strSQL, objConn)
      objCmd.ExecuteNonQuery
      objConn.Close()
  End Sub

  Sub DataList_DeleteCommand(sender As Object, e As DataListCommandEventArgs)
    Dim strSQL As String = "Delete From [OrderMenu] Where [" & _
        myDataList.DataKeyField & "]=" & myDataList.DataKeys(e.Item.ItemIndex)
    ExecuteSQL(strSQL)
    myDataList.EditItemIndex = -1
    BindList()
  End Sub

  Sub DataList_UpdateCommand(sender As Object, e As DataListCommandEventArgs)
    Dim strO_Date As String = Ctype(e.Item.FindControl("txtO_Date"), TextBox).Text
    Dim UpdateDate As DateTime = DateTime.Now.Date()
    Dim strSQL As String = "Update [OrderMenu] Set [O-Date]='" & strO_Date & "' " & _
                           "Where [" & myDataList.DataKeyField & "]=" & myDataList.DataKeys(e.Item.ItemIndex)
    ExecuteSQL(strSQL)
    myDataList.EditItemIndex = -1
    BindList()
  End Sub

  Sub myCalendar_DayRender(sender As Object, e As DayRenderEventArgs)

      Dim d as CalendarDay
      Dim c as TableCell

      d = e.Day
      c = e.Cell

      If d.IsOtherMonth Then
         c.Controls.Clear
      Else
          Try
              Dim Hol As String = holidays(d.Date.Month,d.Date.Day)

              If Hol <> "" Then
                  c.Controls.Add(new LiteralControl("<br>" + Hol))
              End If
          Catch exc as Exception
              Response.Write (exc.ToString())
          End Try
      End If
  End Sub

</SCRIPT>
</HEAD>
  <Body  MS_POSITIONING="GridLayout" background=./image/admin-bg.gif>
    <form id=Form1 method=post runat="server">
      <Asp:DataList Runat="Server" Id="myDataList" CellPadding="3" Width="600px"
       HorizontalAlign="Center" OnEditCommand="DataList_EditCommand"
       OnUpdateCommand="DataList_UpdateCommand" OnDeleteCommand="DataList_DeleteCommand"
       OnCancelCommand="DataList_CancelCommand" DataKeyField="O-ID"
       ExtractTemplateRows="True" BorderColor="#DEBA84" GridLines="Both" RepeatLayout="Flow" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 328px" BackColor="#DEBA84" BorderWidth="1px" BorderStyle="None" CellSpacing="2">
<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#738A9C">
</SelectedItemStyle>

<HeaderTemplate>
          <Asp:Table Runat="Server" Width="600">
            <Asp:TableRow Runat="Server">
              <Asp:TableCell Runat="Server" Width="40" Font-Size ="10pt">
                自動編號
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                日期
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                專業教室
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                節次
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                教師
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                使用內容
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                借用單位
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                功能
              </Asp:TableCell>
            </Asp:TableRow>
          </Asp:Table>
        
</HeaderTemplate>

<EditItemStyle BackColor="Lavender">
</EditItemStyle>

<ItemStyle ForeColor="#8C4510" BackColor="#FFF7E7">
</ItemStyle>

<ItemTemplate>
          <Asp:Table Runat="Server" Width="600">
            <Asp:TableRow Runat="Server" HorizontalAlign="Center">
              <Asp:TableCell Runat="Server" Width="40" Font-Size ="10pt" >
                <%# Container.DataItem("O-ID") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                <%# DateValue(Container.DataItem("O-Date")).ToShortDateString() %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                <%# Container.DataItem("P-Adds") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                <%# Container.DataItem("O-Time") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                <%# Container.DataItem("T-Name_FN_CHT") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                <%# Container.DataItem("B-Name") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                <%# Container.DataItem("C-Name") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                <!--<Asp:LinkButton Runat="Server" Text="編輯" CommandName="Edit"/>-->
              </Asp:TableCell>
            </Asp:TableRow>
          </Asp:Table>
        
</ItemTemplate>

<FooterStyle ForeColor="#8C4510" BackColor="#F7DFB5">
</FooterStyle>

<HeaderStyle Font-Bold="True" HorizontalAlign="Center" ForeColor="White" BackColor="#A55129">
</HeaderStyle>

<EditItemTemplate>
          <Asp:Table Runat="Server" Width="400">
            <Asp:TableRow Runat="Server" HorizontalAlign="Center">
              <Asp:TableCell Runat="Server" Width="70">
                <%# Container.DataItem("O-ID") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="200">
                <Asp:TextBox Runat="Server" Id="txtBook" Width="200" Text='<%# Container.DataItem("O-Date") %>' />
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="130">
                <Asp:LinkButton Runat="Server" Width="18" Text="更新" CommandName="Update" /> 
              　<Asp:LinkButton Runat="Server" Width="18" Text="刪除" CommandName="Delete" /> 
              　<Asp:LinkButton Runat="Server" Width="18" Text="取消" CommandName="Cancel" />
              </Asp:TableCell>
            </Asp:TableRow>
          </Asp:Table>
        
</EditItemTemplate>
      </Asp:DataList><asp:calendar id="myCalendar" style="Z-INDEX: 106; LEFT: 8px; POSITION: absolute; TOP: 60px" runat="server" BorderWidth="1px" BackColor="#FFFFCC" BorderColor="#FFCC66" Width="600px" Height="265px" Font-Size="8pt" ForeColor="#663399" SelectionMode="DayWeekMonth" ShowGridLines="True" OnDayRender="myCalendar_DayRender" OnSelectionChanged="DayChange" Font-Names="Verdana" NextPrevFormat="FullMonth" SelectWeekText="<img src='./image/Week.ico' border='0'></img>" SelectMonthText="<img src='./image/Month.ico' border='0'></img>">
<TodayDayStyle ForeColor="White" BackColor="#FFCC66">
</TodayDayStyle>

<SelectorStyle BackColor="#FFCC66">
</SelectorStyle>

<NextPrevStyle Font-Size="9pt" ForeColor="#FFFFCC">
</NextPrevStyle>

<DayHeaderStyle Height="1px" BackColor="#FFCC66">
</DayHeaderStyle>

<SelectedDayStyle Font-Bold="True" BackColor="#CCCCFF">
</SelectedDayStyle>

<TitleStyle Font-Size="9pt" Font-Bold="True" ForeColor="#FFFFCC" BackColor="#990000">
</TitleStyle>

<OtherMonthDayStyle ForeColor="#CC9966">
</OtherMonthDayStyle>
             </asp:calendar>
      <asp:Label id="Label1" style="Z-INDEX: 105; LEFT: 208px; POSITION: absolute; TOP: 12px" 
                 Width="176px" runat="server" Font-Bold="True" Height="36px" Font-Size="X-Large">預約紀錄表</asp:Label>
    </form>
  </BODY>
</HTML>
