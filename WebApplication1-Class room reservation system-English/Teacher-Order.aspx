<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- #include File="OleDbFunction.inc" --><HTML><HEAD>
		<title>Teacher</title>
		<SCRIPT language="VB" Runat="Server">

    Dim holidays(12,31) as String

				Sub Page_Load( sender As Object, e As Eventargs )
        If Session("LV") <> "Teacher" Then
           Page.Response.Redirect ("./Main.aspx")
           Exit Sub
        End If
        Load_Holidays()
				  	 If Not IsPostBack Then BindList_PS()
				End Sub
				
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


				Sub BindList_PS()
					
						myCalendar.SelectedDate = DateString()
	          
						labMessage.Text = ""

      If Request.QueryString("ID") = "" Then
         Exit Sub
      End If 
					
						Dim strSQL As String 
						Dim objCnn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("./DB/order.mdb") )
						objCnn.Open()
						strSQL = "Select [T-ID],[T-Name_FN_CHT] From [Teacher] Where [T-ID] Like '" & Replace(Request.QueryString("ID"),"'","") & "'"
					     
						Dim objCmd As New OleDbCommand(strSQL,objCnn)
						Dim objReader As OleDbDataReader = objCmd.ExecuteReader()
						If objReader.Read() = True Then
    					labT_Name.Text = "Teacher name¡@" & objReader.Item(1)
						End If
						objReader.Close()
						objCnn.Close()
					  
 						strSQL = "Select [P-ID],[P-Adds] From [PhotoStudio] " 
						rblPS.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "PhotoStudio")
						rblPS.DataTextField="P-Adds"
						rblPS.DataValueField="P-ID"
						rblPS.DataBind()
						rblPS.SelectedIndex = 0
						
 						strSQL = "Select [B-ID],[B-Name] From [Book] " 
						rblBook.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "Book")
						rblBook.DataTextField="B-Name"
						rblBook.DataValueField="B-ID"
						rblBook.DataBind()
						rblBook.SelectedIndex = 0

 						strSQL = "Select [C-ID],[C-Name] From [Class] " 
						rblClass.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "Class")
						rblClass.DataTextField="C-Name"
						rblClass.DataValueField="C-ID"
						rblClass.DataBind()
						rblClass.SelectedIndex = 0

 						strSQL = "Select [I-ID],[I-Name] From [Item] " 
						rblItem.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "Item")
						rblItem.DataTextField="I-Name"
						rblItem.DataValueField="I-ID"
						rblItem.DataBind()
						rblItem.SelectedIndex = 0
						
						objCnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("./DB/order.mdb") 
						objCnn.Open()
'						strSQL = "Select [OrderMenu].[O-Date],[OrderMenu].[O-P_ID],[OrderMenu].[O-Time],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name] From [OrderMenu],[Teacher],[Book],[Class],[PhotoStudio] " & _
'     										"Where [OrderMenu].[O-Date] Like #" & DateString() & "# AND [OrderMenu].[O-P_ID] = " & Clng(rblPS.SelectedItem.Value) & " AND ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] AND [OrderMenu].[O-B_ID] = [Book].[B-ID] AND [OrderMenu].[O-C_ID] = [Class].[C-ID]  AND [OrderMenu].[O-P_ID] = [PhotoStudio].[P-ID] )"
						  strSQL = "Select [OrderMenu].[O-Date],[OrderMenu].[O-P_ID],[OrderMenu].[O-Time],[OrderMenu].[O-I_ID],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name],[Item].[I-Name] From [OrderMenu],[Teacher],[Book],[Class],[PhotoStudio],[Item] " & _
						      			  "Where [OrderMenu].[O-Date] Like #" & DateValue(myCalendar.SelectedDate) & "# AND [OrderMenu].[O-P_ID] = " & Clng(rblPS.SelectedItem.Value) & " AND ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] AND [OrderMenu].[O-B_ID] = [Book].[B-ID] AND [OrderMenu].[O-C_ID] = [Class].[C-ID] AND [OrderMenu].[O-P_ID] = [PhotoStudio].[P-ID] AND [OrderMenu].[O-I_ID] = [Item].[I-ID] )"
			      
						objCmd.Connection = objCnn
						objCmd.CommandText = strSQL
						objReader = objCmd.ExecuteReader()
						While objReader.Read()
								Select Case objReader.Item("O-Time")
											Case 1
		                    rbnTime1.Text = "No. 1:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
														rbnTime1.Enabled = False
											Case 2 
		                    rbnTime2.Text = "No. 2:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
														rbnTime2.Enabled = False
											Case 3
		                    rbnTime3.Text = "No. 3:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
														rbnTime3.Enabled = False
											Case 4
		                    rbnTime4.Text = "No. 4:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
														rbnTime4.Enabled = False
											Case 5
		                    rbnTime5.Text = "No. 5:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
														rbnTime5.Enabled = False
											Case 6
		                    rbnTime6.Text = "No. 6:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
														rbnTime6.Enabled = False
											Case 7
		                    rbnTime7.Text = "No. 7:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
														rbnTime7.Enabled = False
											Case 8
		                    rbnTime8.Text = "No. 8:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
														rbnTime8.Enabled = False
								End Select
						End While
						objReader.Close()
						objCnn.Close()
					
				End Sub
				

				Sub BindList2()

						  labMessage.Text = ""

		        rbnTime1.Text = "No. 1:"
		        rbnTime2.Text = "No. 2:"
		        rbnTime3.Text = "No. 3:"
		        rbnTime4.Text = "No. 4:"
		        rbnTime5.Text = "No. 5:"
		        rbnTime6.Text = "No. 6:"
		        rbnTime7.Text = "No. 7:"
		        rbnTime8.Text = "No. 8:"

						  rbnTime1.Checked = False
						  rbnTime2.Checked = False
						  rbnTime3.Checked = False
						  rbnTime4.Checked = False
						  rbnTime5.Checked = False
						  rbnTime6.Checked = False
						  rbnTime7.Checked = False
						  rbnTime8.Checked = False
  	          
						  rbnTime1.Enabled = True
						  rbnTime2.Enabled = True
						  rbnTime3.Enabled = True
						  rbnTime4.Enabled = True
						  rbnTime5.Enabled = True
						  rbnTime6.Enabled = True
						  rbnTime7.Enabled = True
						  rbnTime8.Enabled = True
  					
						  Dim strSQL As String 
						  Dim objCnn As New OleDbConnection
						  Dim objCmd As New OleDbCommand()
						  Dim objReader As OleDbDataReader 
  					  
						  objCnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                  "Data Source=" & Server.MapPath("./DB/order.mdb") 
						  objCnn.Open()
						  If myCalendar.SelectedDate = Nothing Then
  				 		  myCalendar.SelectedDate = DateString()
						  End If
						  strSQL = "Select [OrderMenu].[O-Date],[OrderMenu].[O-P_ID],[OrderMenu].[O-Time],[OrderMenu].[O-I_ID],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name],[Item].[I-Name] From [OrderMenu],[Teacher],[Book],[Class],[PhotoStudio],[Item] " & _
						      			  "Where [OrderMenu].[O-Date] Like #" & DateValue(myCalendar.SelectedDate) & "# AND [OrderMenu].[O-P_ID] = " & Clng(rblPS.SelectedItem.Value) & " AND ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] AND [OrderMenu].[O-B_ID] = [Book].[B-ID] AND [OrderMenu].[O-C_ID] = [Class].[C-ID] AND [OrderMenu].[O-P_ID] = [PhotoStudio].[P-ID] AND [OrderMenu].[O-I_ID] = [Item].[I-ID] )"
'						  strSQL = "Select [OrderMenu].[O-Date],[OrderMenu].[O-P_ID],[OrderMenu].[O-Time],[OrderMenu].[O-I_ID],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name],[Item].[I-Name] From [OrderMenu],[Teacher],[Book],[Class],[PhotoStudio] " & _
'						      			  "Where [OrderMenu].[O-Date] Like #" & DateValue(myCalendar.SelectedDate) & "# AND [OrderMenu].[O-P_ID] = " & Clng(rblPS.SelectedItem.Value) & " AND ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] AND [OrderMenu].[O-B_ID] = [Book].[B-ID] AND [OrderMenu].[O-C_ID] = [Class].[C-ID] AND [OrderMenu].[O-P_ID] = [PhotoStudio].[P-ID] )"
  			      
						  objCmd.Connection = objCnn
						  objCmd.CommandText = strSQL
						  objReader = objCmd.ExecuteReader()
  			      
						  While objReader.Read()
								    Select Case objReader.Item("O-Time")
											        Case 1
		                    rbnTime1.Text = "No. 1:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
											  			        rbnTime1.Enabled = False
											        Case 2 
		                    rbnTime2.Text = "No. 2:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
 													            rbnTime2.Enabled = False
											        Case 3
		                    rbnTime3.Text = "No. 3:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
		  												        rbnTime3.Enabled = False
											        Case 4
		                    rbnTime4.Text = "No. 4:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
						  								        rbnTime4.Enabled = False
											        Case 5
		                    rbnTime5.Text = "No. 5:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
										  				        rbnTime5.Enabled = False
											        Case 6
		                    rbnTime6.Text = "No. 6:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
		  												        rbnTime6.Enabled = False
											        Case 7
		                    rbnTime7.Text = "No. 7:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
						  								        rbnTime7.Enabled = False
											        Case 8
		                    rbnTime8.Text = "No. 8:" & objReader.Item("T-Name_FN_CHT") & Chr(32) & objReader.Item("B-Name") & Chr(32) & Chr(32) & objReader.Item("C-Name") & Chr(32) & Chr(32) & objReader.Item("I-Name")
										  				        rbnTime8.Enabled = False
								    End Select
						  End While
		      
  				  objReader.Close()
	  			  objCnn.Close()

				End Sub
				
				Sub DayChange( sender As Object, e As Eventargs )
  				  Call BindList2()
				End Sub

				Sub PSChange( sender As Object, e As Eventargs )
		  		  Call BindList2()
				End Sub
				
				Sub OrderClass( sender As Object, e As Eventargs )

        If IsDate(myCalendar.SelectedDate) = False Then
		            labMessage.Text = "Please, choice a date."
           Exit Sub
        End If

        If myCalendar.SelectedDate < DateString() Then
		            labMessage.Text = "Reservation date is error."
           Exit Sub
        End If
 
				    Dim strO_Time as String = "0"
  				  
				    If rbnTime1.Checked = True Then strO_Time = "1"
				    If rbnTime2.Checked = True Then strO_Time = "2"
				    If rbnTime3.Checked = True Then strO_Time = "3"
				    If rbnTime4.Checked = True Then strO_Time = "4"
				    If rbnTime5.Checked = True Then strO_Time = "5"
				    If rbnTime6.Checked = True Then strO_Time = "6"
				    If rbnTime7.Checked = True Then strO_Time = "7"
				    If rbnTime8.Checked = True Then strO_Time = "8"

        If strO_Time = "0" Then
		            labMessage.Text = "Please, select a time."
           Exit Sub
        End If
  				  
					   Dim objConn As New OleDbConnection()
					   objConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
																	                  "Data Source=" & Server.MapPath("./DB/order.mdb")
					   objConn.Open()

        Dim objCmd As New OleDbCommand("Select [O-Date],[O-P_ID],[O-Time] From [OrderMenu] " & _
                                       "Where [O-Date] Like #" & DateValue(myCalendar.SelectedDate) & "# And [O-P_ID] = " & (rblPS.SelectedItem.Value) & " And [O-Time] Like '" & strO_Time & "'" , _
                                       objConn) 
        Dim objReader As OleDbDataReader = objCmd.ExecuteReader()
        
        If objReader.Read() = True Then
		            labMessage.Text = "It had been reservation."
           objConn.Close()
           Exit Sub
        End If
        objReader.Close
        objCmd.Dispose

		        'Dim objCmd As New OleDbCommand("Insert Into OrderMenu ([O-Date], [O-P_ID], [O-Time], [O-T_ID], [O-B_ID], [O-C_ID]) Values ('" & DateString() & "','" & (rblPS.SelectedItem.Value) & "','" & strO_Time & "','" & Request.QueryString("ID") & "','" & (rblBook.SelectedItem.Value) & "','" & (rblClass.SelectedItem.Value) & "')", objConn)
					   
        If CheckBox1.checked = True Then
           If Len(Trim(TextBox1.Text & "")) > 0 Then
				objCmd.Connection = objConn
				objCmd.CommandText = "Insert Into PhotoStudio ([P-Adds]) " & _
									"Values ('" & Trim(TextBox1.Text & "") & "')"
							objCmd.ExecuteNonQuery()
              
           End If
        End If

        If CheckBox2.checked = True Then
           If Len(Trim(TextBox2.Text & "")) > 0 Then
				objCmd.Connection = objConn
				objCmd.CommandText = "Insert Into Book ([B-Name]) " & _
									"Values ('" & Trim(TextBox2.Text & "") & "')"
							objCmd.ExecuteNonQuery()
              
           End If
        End If

        If CheckBox3.checked = True Then
           If Len(Trim(TextBox3.Text & "")) > 0 Then
				objCmd.Connection = objConn
				objCmd.CommandText = "Insert Into Class ([C-Name]) " & _
									"Values ('" & Trim(TextBox3.Text & "") & "')"
							objCmd.ExecuteNonQuery()
              
           End If
        End If

        If CheckBox4.checked = True Then
           If Len(Trim(TextBox4.Text & "")) > 0 Then
				objCmd.Connection = objConn
				objCmd.CommandText = "Insert Into Item ([I-Name]) " & _
									"Values ('" & Trim(TextBox4.Text & "") & "')"
							objCmd.ExecuteNonQuery()
              
           End If
        End If
					   
        Dim REC_PS , REC_BOOK , REC_CLASS , REC_Item as integer 
        
        REC_PS = rblPS.SelectedItem.Value
        REC_BOOK = rblBook.SelectedItem.Value
        REC_CLASS = rblClass.SelectedItem.Value
        REC_Item = rblItem.SelectedItem.Value


		        Dim objCmd_PS As New OleDbCommand("Select [P-ID],[P-Adds] From [PhotoStudio] " & _
		                               "Where [P-Adds] Like '" & Trim(TextBox1.Text & "") & "'", _
		                               objConn)
        Dim objReader_PS As OleDbDataReader = objCmd_PS.ExecuteReader()
        
        If objReader_PS.Read() = True Then
           REC_PS = objReader_PS.Item("P-ID")
        End If
        objReader_PS.Close
        objCmd_PS.Dispose


		        'Context
        Dim objCmd_BOOK As New OleDbCommand("Select [B-ID],[B-Name] From [Book] " & _
                                       "Where [B-Name] Like '" & Trim(TextBox2.Text & "") & "'" , _
                                       objConn) 
        Dim objReader_BOOK As OleDbDataReader = objCmd_BOOK.ExecuteReader()
        
        If objReader_BOOK.Read() = True Then
           REC_BOOK = objReader_BOOK.Item("B-ID")
        End If
        objReader_BOOK.Close
        objCmd_BOOK.Dispose
        
        
		        'Office
        Dim objCmd_CLASS As New OleDbCommand("Select [C-ID],[C-Name] From [Class] " & _
                                       "Where [C-Name] Like '" & Trim(TextBox3.Text & "") & "'" , _
                                       objConn) 
        Dim objReader_CLASS As OleDbDataReader = objCmd_CLASS.ExecuteReader()
        
        If objReader_CLASS.Read() = True Then
           REC_CLASS = objReader_CLASS.Item("C-ID")
        End If
        objReader_CLASS.Close
        objCmd_CLASS.Dispose
        

		        'Item
        Dim objCmd_Item As New OleDbCommand("Select [I-ID],[I-Name] From [Item] " & _
                                       "Where [I-Name] Like '" & Trim(TextBox4.Text & "") & "'" , _
                                       objConn) 
        Dim objReader_Item As OleDbDataReader = objCmd_Item.ExecuteReader()
        
        If objReader_Item.Read() = True Then
           REC_Item = objReader_Item.Item("I-ID")
        End If
        objReader_Item.Close
        objCmd_Item.Dispose
					   
        objCmd.Connection = objConn
        objCmd.CommandText = "Insert Into OrderMenu ([O-Date], [O-P_ID], [O-Time], [O-T_ID], [O-B_ID], [O-C_ID], [O-I_ID]) " & _
                             "Values ('" & DateValue(myCalendar.SelectedDate) & "','" & (REC_PS) & "','" & strO_Time & "','" & Request.QueryString("ID") & "','" & (REC_BOOK) & "','" & (REC_CLASS) & "','" & (REC_Item) & "')"
					   objCmd.ExecuteNonQuery()
					   objConn.Close()
      				  
				    Call BindList2()

		        labMessage.Text = "Reservation successful!"

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
		<meta content="Microsoft Visual Studio .NET 10.0" name="GENERATOR">
		<meta content="Visual Basic .NET 10.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body background="./image/Teacher-bg.gif" MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="Arial">
				<asp:label id="LabPS" style="Z-INDEX: 101; LEFT: 12px; POSITION: absolute; TOP: 344px" runat="server"
					Height="20px" Width="80px">Classroom</asp:label>
				<asp:TextBox id="TextBox4" style="Z-INDEX: 131; LEFT: 172px; POSITION: absolute; TOP: 680px"
					runat="server" Height="20px" Width="188px"></asp:TextBox>
				<asp:CheckBox id="CheckBox4" style="Z-INDEX: 130; LEFT: 104px; POSITION: absolute; TOP: 680px"
					runat="server" Height="16px" Width="64px" Text="Other" Font-Size="10pt"></asp:CheckBox>
				<asp:radiobuttonlist id="rblItem" style="Z-INDEX: 129; LEFT: 104px; POSITION: absolute; TOP: 708px" runat="server"
					Height="80px" Width="488px" Font-Size="10pt" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:radiobuttonlist>
				<asp:label id="Label5" style="Z-INDEX: 128; LEFT: 12px; POSITION: absolute; TOP: 676px" runat="server"
					Height="20px" Width="80px">Item</asp:label>
				<asp:TextBox id="TextBox3" style="Z-INDEX: 127; LEFT: 172px; POSITION: absolute; TOP: 564px"
					runat="server" Height="20px" Width="188px"></asp:TextBox>
				<asp:CheckBox id="CheckBox3" style="Z-INDEX: 125; LEFT: 104px; POSITION: absolute; TOP: 564px"
					runat="server" Height="16px" Width="64px" Text="Other" Font-Size="10pt"></asp:CheckBox>
				<asp:CheckBox id="CheckBox2" style="Z-INDEX: 123; LEFT: 104px; POSITION: absolute; TOP: 452px"
					runat="server" Height="16px" Width="64px" Text="Other" Font-Size="10pt"></asp:CheckBox><asp:label id="labDate" style="Z-INDEX: 118; LEFT: 12px; POSITION: absolute; TOP: 96px" runat="server"
					Height="20px" Width="80px">Reservation date</asp:label><asp:label id="labT_Name" style="Z-INDEX: 117; LEFT: 12px; POSITION: absolute; TOP: 60px" runat="server"
					Height="20px" Width="340px">Teacher name</asp:label>
                    <asp:radiobutton id="rbnTime8" style="Z-INDEX: 114; LEFT: 488px; POSITION: absolute; TOP: 860px" runat="server" Height="44px" Width="120px" Text="No. 8:" GroupName="gTime" Font-Size="10pt"></asp:radiobutton><br />
                    <asp:radiobutton id="rbnTime7" style="Z-INDEX: 113; LEFT: 360px; POSITION: absolute; TOP: 860px" runat="server" Height="44px" Width="120px" Text="No. 7:" GroupName="gTime" Font-Size="10pt"></asp:radiobutton><br />
                    <asp:radiobutton id="rbnTime6" style="Z-INDEX: 112; LEFT: 232px; POSITION: absolute; TOP: 860px" runat="server" Height="44px" Width="120px" Text="No. 6:" GroupName="gTime" Font-Size="10pt"></asp:radiobutton><br />
                    <asp:radiobutton id="rbnTime5" style="Z-INDEX: 111; LEFT: 104px; POSITION: absolute; TOP: 860px" runat="server" Height="44px" Width="120px" Text="No. 5:" GroupName="gTime" Font-Size="10pt"></asp:radiobutton><br />
                    <asp:radiobutton id="rbnTime4" style="Z-INDEX: 110; LEFT: 488px; POSITION: absolute; TOP: 792px" runat="server" Height="44px" Width="120px" Text="No. 4:" GroupName="gTime" Font-Size="10pt"></asp:radiobutton><br />
                    <asp:radiobutton id="rbnTime3" style="Z-INDEX: 109; LEFT: 360px; POSITION: absolute; TOP: 792px" runat="server" Height="44px" Width="120px" Text="No. 3:" GroupName="gTime" Font-Size="10pt"></asp:radiobutton><br />
                    <asp:radiobutton id="rbnTime2" style="Z-INDEX: 108; LEFT: 232px; POSITION: absolute; TOP: 792px" runat="server" Height="44px" Width="120px" Text="No. 2:" GroupName="gTime" Font-Size="10pt"></asp:radiobutton><br />
                    <asp:radiobutton id="rbnTime1" style="Z-INDEX: 107; LEFT: 104px; POSITION: absolute; TOP: 792px" runat="server" Height="44px" Width="120px" Text="No. 1:" GroupName="gTime" Font-Size="10pt"></asp:radiobutton>
                    <asp:label id="Label3" style="Z-INDEX: 106; LEFT: 12px; POSITION: absolute; TOP: 800px" runat="server"
					Height="20px" Width="80px">Time</asp:label><asp:radiobuttonlist id="rblClass" style="Z-INDEX: 105; LEFT: 104px; POSITION: absolute; TOP: 592px"
					runat="server" Height="80px" Width="488px" Font-Size="10pt" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:radiobuttonlist><asp:label id="Label2" style="Z-INDEX: 104; LEFT: 12px; POSITION: absolute; TOP: 568px" runat="server"
					Height="20px" Width="80px">Office</asp:label><asp:radiobuttonlist id="rblBook" style="Z-INDEX: 103; LEFT: 104px; POSITION: absolute; TOP: 480px" runat="server"
					Height="80px" Width="488px" Font-Size="10pt" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:radiobuttonlist><asp:label id="Label1" style="Z-INDEX: 102; LEFT: 12px; POSITION: absolute; TOP: 456px" runat="server"
					Height="20px" Width="80px">Context</asp:label><asp:radiobuttonlist id="rblPS" style="Z-INDEX: 100; LEFT: 104px; POSITION: absolute; TOP: 368px" runat="server"
					Height="80px" Width="488px" Font-Size="10pt" RepeatDirection="Horizontal" RepeatLayout="Flow" OnSelectedIndexChanged="PSChange" AutoPostBack="True"></asp:radiobuttonlist><asp:label id="Label4" style="Z-INDEX: 115; LEFT: 156px; POSITION: absolute; TOP: 8px" runat="server"
					Height="36px" Width="368px" Font-Size="X-Large">Teacher reservation application</asp:label><asp:button id="myButton" style="Z-INDEX: 116; LEFT: 12px; POSITION: absolute; TOP: 888px" onclick="OrderClass"
					runat="server" Width="80px" Text="Reservation"></asp:button><asp:calendar id="myCalendar" style="Z-INDEX: 119; LEFT: 104px; POSITION: absolute; TOP: 84px"
					runat="server" Height="252px" Width="492px" Font-Size="10pt" OnDayRender="myCalendar_DayRender" NextPrevFormat="FullMonth" BackColor="White" ForeColor="Black"
					Font-Names="Times New Roman" BorderColor="Black" OnSelectionChanged="DayChange">
					<TodayDayStyle BackColor="#CCCC99"></TodayDayStyle>
					<SelectorStyle Font-Size="8pt" Font-Names="Verdana" Font-Bold="True" ForeColor="#333333" Width="1%"
						BackColor="#CCCCCC"></SelectorStyle>
					<DayStyle Width="14%"></DayStyle>
					<NextPrevStyle Font-Size="8pt" ForeColor="White"></NextPrevStyle>
					<DayHeaderStyle Font-Size="7pt" Font-Names="Verdana" Font-Bold="True" Height="10px" ForeColor="#333333"
						BackColor="#CCCCCC"></DayHeaderStyle>
					<SelectedDayStyle ForeColor="White" BackColor="#CC3333"></SelectedDayStyle>
					<TitleStyle Font-Size="13pt" Font-Bold="True" Height="14pt" ForeColor="White" BackColor="Black"></TitleStyle>
					<OtherMonthDayStyle ForeColor="#999999"></OtherMonthDayStyle>
				</asp:calendar><asp:label id="labMessage" style="Z-INDEX: 120; LEFT: 104px; POSITION: absolute; TOP: 892px"
					runat="server" Height="20px" Width="488px"></asp:label>
				<asp:TextBox id="TextBox1" style="Z-INDEX: 121; LEFT: 172px; POSITION: absolute; TOP: 340px"
					runat="server" Height="20px" Width="184px" Font-Size="10pt"></asp:TextBox>
				<asp:CheckBox id="CheckBox1" style="Z-INDEX: 122; LEFT: 104px; POSITION: absolute; TOP: 340px"
					runat="server" Height="16px" Width="64px" Text="Other" Font-Size="10pt"></asp:CheckBox>
				<asp:TextBox id="TextBox2" style="Z-INDEX: 124; LEFT: 172px; POSITION: absolute; TOP: 452px"
					runat="server" Height="20px" Width="188px"></asp:TextBox></FONT></form>
	</body>
</HTML>
