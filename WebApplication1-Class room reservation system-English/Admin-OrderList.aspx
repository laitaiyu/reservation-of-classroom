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
        holidays(1, 1) = "*New Year's Day"
        holidays(2, 12) = "Abraham Lincoln's Birthday"
        holidays(2, 14) = "St. Valentine's Day"
        holidays(2, 22) = "George Washington's Birthday"
        holidays(3, 17) = "St. Patrick's Day"
        holidays(3, 7) = "Easter Day, Easter Sunday"
        holidays(4, 1) = "April Fool's Day"
        holidays(5, 10) = "Mother's Day"
        holidays(5, 25) = "Memorial Day"
        holidays(6, 14) = "National Flag Day"
        holidays(6, 21) = "Father's Day"
        holidays(7, 4) = "Independence Day"
        holidays(9, 6) = "Labor Day"
        holidays(10, 12) = "Columbus Day"
        holidays(10, 31) = "Halloween�FEve of All Saint's Day"
        holidays(11, 2) = "All Soul's Day"
        holidays(11, 11) = "Veterans Day"
        holidays(11, 26) = "Thanksgiving Day"
        holidays(12, 21) = "Forefather's Day"
        holidays(12, 25) = "Christmas Day"
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
                ID
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                Date
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                Professional Classroom
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                Time
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                Teacher
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                Context
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                Office
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                Function
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
                <!--<Asp:LinkButton Runat="Server" Text="Edit" CommandName="Edit"/>-->
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
                <Asp:LinkButton Runat="Server" Width="18" Text="Update" CommandName="Update" /> 
               , <Asp:LinkButton Runat="Server" Width="18" Text="Delete" CommandName="Delete" /> 
               , <Asp:LinkButton Runat="Server" Width="18" Text="Cancel" CommandName="Cancel" />
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
                 Width="176px" runat="server" Font-Bold="True" Height="36px" Font-Size="X-Large">Reservation List</asp:Label>
    </form>
  </BODY>
</HTML>
