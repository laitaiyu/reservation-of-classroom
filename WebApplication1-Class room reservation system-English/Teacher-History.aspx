<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Teacher reservation list</title> 
		<!-- #include File="OleDbFunction.inc" -->
		<SCRIPT language="VB" Runat="Server">

    Dim holidays(12,31) as String

    Dim strT_ID As String 

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
		        holidays(10, 31) = "Halloween¡FEve of All Saint's Day"
		        holidays(11, 2) = "All Soul's Day"
		        holidays(11, 11) = "Veterans Day"
		        holidays(11, 26) = "Thanksgiving Day"
		        holidays(12, 21) = "Forefather's Day"
		        holidays(12, 25) = "Christmas Day"
		    End Sub

				Sub Page_Load( sender As Object, e As Eventargs )
        If Session("LV") <> "Teacher" Then
           Page.Response.Redirect ("./Main.aspx")
           Exit Sub
        End If
        strT_ID = Request.QueryString("ID")
        Load_Holidays()
		  	 		If Not IsPostBack Then BindList_PS()
		  	 		If Not IsPostBack Then BindList()
				End Sub

				Sub BindList_PS()
  				  myCalendar.SelectedDate = DateString()

        If Request.QueryString("ID") = "" Then
           Exit Sub
        End If 
   				
						  Dim strSQL As String 
		        Dim objCnn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("./DB/order.mdb"))
		        objCnn.Open()
		        strSQL = "Select [T-ID],[T-Name_FN_CHT] From [Teacher] Where [T-ID] Like '" & Request.QueryString("ID") & "'"
   					    
		        Dim objCmd As New OleDbCommand(strSQL, objCnn)
		        Dim objReader As OleDbDataReader = objCmd.ExecuteReader()
		        If objReader.Read() = True Then
		            labT_Name.Text = "Teacher name " & objReader.Item(1)
		        End If
		        objReader.Close()
		        objCnn.Close()


		        strSQL = "Select [P-ID],[P-Adds] From [PhotoStudio] "
		        myRadioButtonList.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "PhotoStudio")
		        myRadioButtonList.DataTextField = "P-Adds"
		        myRadioButtonList.DataValueField = "P-ID"
		        myRadioButtonList.DataBind()
		        myRadioButtonList.SelectedIndex = 0
		    End Sub

		    Sub BindList()
		        Dim strSQL As String
		        strSQL = "Select [OrderMenu].[O-T_ID],[OrderMenu].[O-I_ID],[OrderMenu].[O-B_ID],[OrderMenu].[O-C_ID],[OrderMenu].[O-P_ID],[OrderMenu].[O-Date],[OrderMenu].[O-Time],[Teacher].[T-Name_LN_CHT],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name],[Item].[I-Name] From [OrderMenu],[Teacher],[Book],[Class],[Item] " & _
		                 "Where [OrderMenu].[O-Date] Like #" & DateString() & "# And [O-T_ID] =" & CLng(strT_ID) & " And [OrderMenu].[O-P_ID] = " & CLng(myRadioButtonList.SelectedItem.Value) & " And ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] And [OrderMenu].[O-B_ID] = [Book].[B-ID] And [OrderMenu].[O-C_ID] = [Class].[C-ID] And [OrderMenu].[O-I_ID] = [Item].[I-ID]) Order By [OrderMenu].[O-Date],[OrderMenu].[O-Time]"
		        myDataGrid.CurrentPageIndex = 0
		        myDataGrid.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "OrderMenu")
		        myDataGrid.DataBind()
		    End Sub

		    Sub BindList2()
		        Dim strSQL As String
		        If myCalendar.SelectedDates.Count = 1 Then
		            strSQL = "Select [OrderMenu].[O-T_ID],[OrderMenu].[O-I_ID],[OrderMenu].[O-B_ID],[OrderMenu].[O-C_ID],[OrderMenu].[O-P_ID],[OrderMenu].[O-Date],[OrderMenu].[O-Time],[Teacher].[T-Name_LN_CHT],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name],[Item].[I-Name] From [OrderMenu],[Teacher],[Book],[Class],[Item]  " & _
		                     "Where [OrderMenu].[O-Date] Like #" & DateValue(myCalendar.SelectedDate) & "#" & " And [O-P_ID] = " & CLng(myRadioButtonList.SelectedItem.Value) & " And [O-T_ID] =" & CLng(strT_ID) & " And ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] And [OrderMenu].[O-B_ID] = [Book].[B-ID] And [OrderMenu].[O-C_ID] = [Class].[C-ID] And [OrderMenu].[O-I_ID] = [Item].[I-ID]) Order By [OrderMenu].[O-Date],[OrderMenu].[O-Time]"
		            myDataGrid.CurrentPageIndex = 0

		        ElseIf myCalendar.SelectedDates.Count > 1 Then
		            Dim strFirstDate As String
		            Dim strLastDate As String
		            With myCalendar.SelectedDates
		                strFirstDate = .Item(0)
		                strLastDate = .Item(.Count - 1)
		            End With
		            strSQL = "Select [OrderMenu].[O-T_ID], [OrderMenu].[O-I_ID],[OrderMenu].[O-B_ID],[OrderMenu].[O-C_ID],[OrderMenu].[O-P_ID],[OrderMenu].[O-Date],[OrderMenu].[O-Time],[Teacher].[T-Name_LN_CHT],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name],[Item].[I-Name] From [OrderMenu],[Teacher],[Book],[Class],[Item] " & _
		                     "Where [OrderMenu].[O-Date] Between #" & strFirstDate & "# And #" & strLastDate & "# And [O-P_ID] = " & CLng(myRadioButtonList.SelectedItem.Value) & " And [O-T_ID] =" & CLng(strT_ID) & " And ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] And [OrderMenu].[O-B_ID] = [Book].[B-ID] And [OrderMenu].[O-C_ID] = [Class].[C-ID] And [OrderMenu].[O-I_ID] = [Item].[I-ID]) Order By [OrderMenu].[O-Date],[OrderMenu].[O-Time]"

		        End If
		        myDataGrid.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "OrderMenu")
		        myDataGrid.DataBind()
		    End Sub

    Sub myDataGrid_PageIndexChanged( sender As Object, e As DataGridPageChangedEventArgs )
        myDataGrid.CurrentPageIndex = e.NewPageIndex
        Call BindList2()
    End Sub

				Sub DayChange( sender As Object, e As Eventargs )
        myDataGrid.CurrentPageIndex = 0 
				    Call BindList2()
				End Sub

				
				Sub PhotoStudioChange ( sender as Object, e As Eventargs )
        myDataGrid.CurrentPageIndex = 0 
				    Call BindList2()
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
		<meta content="True" name="vs_snapToGrid">
		<meta content="True" name="vs_showGrid">
		<meta content="Microsoft Visual Studio .NET 10.0" name="GENERATOR">
		<meta content="Visual Basic .NET 10.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout" background="./image/Teacher-bg.gif">
		<form id="Form1" method="post" runat="server">
				<asp:label id="LabPS" style="Z-INDEX: 106; LEFT: 8px; POSITION: absolute; TOP: 72px" runat="server"
					Height="20px" Width="80px">Classroom </asp:label><asp:label id="labT_Name" style="Z-INDEX: 118; LEFT: 8px; POSITION: absolute; TOP: 48px" runat="server"
					Width="340px" Height="20px">Teacher </asp:label><asp:radiobuttonlist id="myRadioButtonList" style="Z-INDEX: 105; LEFT: 96px; POSITION: absolute; TOP: 72px"
					runat="server" Height="44px" Width="424px" AutoPostBack="True" RepeatLayout="Flow" RepeatDirection="Horizontal" OnSelectedIndexChanged="PhotoStudioChange"></asp:radiobuttonlist><asp:calendar id="myCalendar" style="Z-INDEX: 102; LEFT: 8px; POSITION: absolute; TOP: 120px"
					runat="server" Height="264px" Width="540px" SelectionMode="DayWeekMonth" OnDayRender="myCalendar_DayRender" BorderWidth="1px" OnSelectionChanged="DayChange" ForeColor="#003399" Font-Size="8pt" Font-Names="Verdana" BackColor="White" BorderColor="#3366CC" SelectWeekText="<img src='./image/Week.ico' border='0'></img>"
					SelectMonthText="<img src='./image/Month.ico' border='0'></img>" CellPadding="1" NextPrevFormat="ShortMonth">
					<TodayDayStyle ForeColor="White" BackColor="#99CCCC"></TodayDayStyle>
					<SelectorStyle ForeColor="#336666" BackColor="#99CCCC"></SelectorStyle>
					<NextPrevStyle Font-Size="8pt" ForeColor="#CCCCFF"></NextPrevStyle>
					<DayHeaderStyle Height="1px" ForeColor="#336666" BackColor="#99CCCC"></DayHeaderStyle>
					<SelectedDayStyle Font-Bold="True" ForeColor="#CCFF99" BackColor="#009999"></SelectedDayStyle>
					<TitleStyle Font-Size="10pt" Font-Bold="True" Height="25px" BorderWidth="1px" ForeColor="#CCCCFF"
						BorderStyle="Solid" BorderColor="#3366CC" BackColor="#003399"></TitleStyle>
					<WeekendDayStyle BackColor="#CCCCFF"></WeekendDayStyle>
					<OtherMonthDayStyle ForeColor="#999999"></OtherMonthDayStyle>
				</asp:calendar>
			<DIV style="Z-INDEX: 103; LEFT: 156px; WIDTH: 260px; POSITION: absolute; TOP: 12px; HEIGHT: 25px"
				ms_positioning="FlowLayout"><SPAN style="FONT-SIZE: 16pt; FONT-FAMILY: Arial; mso-bidi-font-size: 12.0pt; mso-hansi-font-family: 'Times New Roman'; mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: EN-US; mso-bidi-language: AR-SA"><STRONG>Teacher reservation list</STRONG>
				</SPAN>
			</DIV>
			<asp:datagrid id="myDataGrid" style="Z-INDEX: 104; LEFT: 8px; POSITION: absolute; TOP: 384px"
				runat="server" Width="540px" BorderWidth="1px" BackColor="White" BorderColor="#E7E7FF" AutoGenerateColumns="False"
				PageSize="8" CellPadding="3" BorderStyle="None" AllowPaging="True" OnPageIndexChanged="myDataGrid_PageIndexChanged"
				GridLines="Horizontal">
				<FooterStyle ForeColor="#4A3C8C" BackColor="#B5C7DE"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#F7F7F7" BackColor="#738A9C"></SelectedItemStyle>
				<AlternatingItemStyle BackColor="#F7F7F7"></AlternatingItemStyle>
				<ItemStyle ForeColor="#4A3C8C" BackColor="#E7E7FF"></ItemStyle>
				<HeaderStyle Font-Size="10pt" Font-Bold="True" ForeColor="#F7F7F7" BackColor="#4A3C8C"></HeaderStyle>
				<Columns>
					<asp:BoundColumn DataField="O-Date" HeaderText="Date" DataFormatString="{0:d}"></asp:BoundColumn>
					<asp:BoundColumn DataField="O-Time" HeaderText="Time"></asp:BoundColumn>
					<asp:BoundColumn DataField="B-Name" HeaderText="Context"></asp:BoundColumn>
					<asp:BoundColumn DataField="C-Name" HeaderText="Office"></asp:BoundColumn>
					<asp:BoundColumn DataField="I-Name" HeaderText="Item"></asp:BoundColumn>
				</Columns>
				<PagerStyle HorizontalAlign="Center" ForeColor="#4A3C8C" BackColor="#E7E7FF" Mode="NumericPages"></PagerStyle>
			</asp:datagrid></form>
	</body>
</HTML>
