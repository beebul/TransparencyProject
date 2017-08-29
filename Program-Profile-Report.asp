<!DOCTYPE html>
<html>

	<div class="logo">
		<img src="uni-logo.jpg" class="heading">
		<h1 class="heading">ADMISSIONS INFORMATION SET</h1>
	</div>

	<%
			Dim conn
			Set conn = Server.CreateObject("ADODB.Connection")
			''conn.Open "Provider=SQLOLEDB;Data Source=.\SQLEXPRESS;Initial Catalog=transparency;User Id=sa;Password=KLF-chill-out-1990"
			conn.Open "Provider=SQLOLEDB;Data Source=.\SQLEXPRESS;Initial Catalog=transparency;User Id=sa;Password=Homerjones1"

			ppc = Request.QueryString("ppc")

			set rs = Server.CreateObject("ADODB.recordset")
			sql = " SELECT * from ProgramProfile WHERE ProgramPlanCode = '" & ppc & "' ORDER BY FINAL_GROUPING"
			rs.Open sql, conn

			satac_code = rs.Fields("SATACCode")

			Dim metatitle, metaDesc, progCode
			metatitle = "<title>" & ppc & "</title>"
			metaDesc = "<meta name=""description"" content=""" & ppc & """ />"
			progCode = rs.Fields("ProgCode")

			Function expandOT(offeringCode)
				Select Case offeringCode
	  			Case "I"
	    			expandOT = "Internal Offering"
	  			Case "P"
	    			expandOT = "Packaged Offering"
	  			Case "E"
	    			expandOT = "External Offering"
				End Select
			End Function

	If (rs.EOF) Then
		Response.Write("<h4 class=""error"">Errors - Query returned no results for Program Plan Code : ''" & ppc & "''</h4>")
	Else

	Response.write("<h3 class=""programName heading""> Program Name : <span>" & rs.Fields("ProgramName") & "</span>&nbsp;/&nbsp;Program Code : <span>" & progCode & "</span></h3>	<br>")
  If (rs.Fields("Offering") <> "") Then
		offeringType = Trim(rs.Fields("Offering"))
		fullOfferingType= expandOT(offeringType)
		Response.Write("<h3 class=""programName heading"" style=""margin-top: -1em;"">" & fullOfferingType &  " / SATAC CODE: " & satac_code & "</h3>")
	End If
	%>

<head>
	<%= metatitle %>
	<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	<meta name="keywords" content="test website, web testing">
	<%= metaDesc %>
	<meta name="robots" content="all">
	<link rel="stylesheet" href="https://unpkg.com/purecss@1.0.0/build/pure-min.css" integrity="sha384-nn4HPE8lTHyVtfCBi5yW9d20FjT8BJwUXyWZT9InLYax14RDjBj46LmSztkmNP9w" crossorigin="anonymous">
	<!-- <link href="//fonts.googleapis.com/css?family=Montserrat" rel="stylesheet"> -->
	<link rel="stylesheet" type="text/css" href="transparency.css">
</head>

<body>

		<blockquote>The admissions information set for each degree outlines a collection of information that will help you to gauge and compare study options, course admission criteria and your likely student peer cohort across multiple courses and admission options. Each admissions information set at the degree level will include a student profile for the degree and, for most degrees, an ATAR profile. An ATAR profile will not be available where there were no applicants in the secondary education applicant group who were admitted solely on the basis of their ATAR for that degree.<br>
		<br><b>Student profile</b><br>
		The table below gives an indication of the likely peer cohort for new students at UniSA. It provides data of students who commenced undergraduate study in the most recent intake period.
		<ul>
			<b>Note:</b>
			<li>L/N - low numbers: the numbers of students is less than 5</li>
			<li>N/A - data not available for this item</li>
			<li>N/P - Not published: the number is hidden to prevent calculation if number in cells with less than 5.</li>
			<li>Group C: UniSA does not admit students where both ATAR and additional criteria are considered. Therefore this Group C subgroup has been omitted from the table.</li>
		</ul>
	</blockquote>
	<br>

			<table class="pure-table pure-table-bordered" width="100%">
				<thead>
				<tr>
			    <th>Applicant Background</th>
			    <th>Number of Students</th>
			    <th>Percentage of all Students</th>
			  </tr>
			</thead>

				<% do until rs.EOF

					finalGroupingA = rs.Fields("ApplicantBackground")
					if Left(finalGroupingA, 3) = "(A)" Then
						Response.Write("<tr><td>" & rs.Fields("ApplicantBackground")) & "</td>"
						Response.Write("<td>" & rs.Fields("NumberOfStudents")) & "</td>"
						Response.Write("<td>" & rs.Fields("PercentageOfAllStudents"))
						if rs.Fields("PercentageOfAllStudents") = "N/P" Then
							Response.Write("</td></tr>")
						Else
							Response.Write("%</td></tr>")
						End If
					End If

					finalGroupingB = rs.Fields("ApplicantBackground")
					if Left(finalGroupingB, 3) = "(B)" Then
						Response.Write("<tr><td>" & rs.Fields("ApplicantBackground")) & "</td>"
						Response.Write("<td>" & rs.Fields("NumberOfStudents")) & "</td>"
						Response.Write("<td>" & rs.Fields("PercentageOfAllStudents"))
						if rs.Fields("PercentageOfAllStudents") = "N/P" Then
							Response.Write("</td></tr>")
						Else
							Response.Write("%</td></tr>")
						End If
					End If

					finalGroupingC = rs.Fields("ApplicantBackground")
					if Left(finalGroupingC, 3) = "(C)" Then
						Response.Write("<tr><td>" & rs.Fields("ApplicantBackground")) & "</td>"
						Response.Write("<td>" & rs.Fields("NumberOfStudents")) & "</td>"
						Response.Write("<td>" & rs.Fields("PercentageOfAllStudents"))
						if rs.Fields("PercentageOfAllStudents") = "N/P" Then
							Response.Write("</td></tr>")
						Else
							Response.Write("%</td></tr>")
						End If
					End If

					finalGroupingD = rs.Fields("ApplicantBackground")
					if Left(finalGroupingD, 3) = "(D)" Then
						Response.Write("<tr><td>" & rs.Fields("ApplicantBackground")) & "</td>"
						Response.Write("<td>" & rs.Fields("NumberOfStudents")) & "</td>"
						Response.Write("<td>" & rs.Fields("PercentageOfAllStudents"))
						if rs.Fields("PercentageOfAllStudents") = "N/P" Then
							Response.Write("</td></tr>")
						Else
							Response.Write("%</td></tr>")
						End If
					End If

					finalGroupingE = rs.Fields("ApplicantBackground")
					if Left(finalGroupingE, 3) = "Int" Then
						Response.Write("<tr><td>" & rs.Fields("ApplicantBackground")) & "</td>"
						Response.Write("<td>" & rs.Fields("NumberOfStudents")) & "</td>"
						Response.Write("<td>" & rs.Fields("PercentageOfAllStudents"))
						if rs.Fields("PercentageOfAllStudents") = "N/P" Then
							Response.Write("</td></tr>")
						Else
							Response.Write("%</td></tr>")
						End If
					End If

					finalGroupingF = rs.Fields("ApplicantBackground")
					if Left(finalGroupingF, 3) = "All" Then
						Response.Write("<tr><td>" & rs.Fields("ApplicantBackground")) & "</td>"
						Response.Write("<td>" & rs.Fields("NumberOfStudents")) & "</td>"
						Response.Write("<td>" & rs.Fields("PercentageOfAllStudents"))
						if rs.Fields("PercentageOfAllStudents") = "N/P" Then
							Response.Write("</td></tr>")
						Else
							Response.Write("%</td></tr>")
						End If
					End If

				rs.MoveNext
				loop %>
			</table>

			<%
			End If

					set rs = Server.CreateObject("ADODB.recordset")

					 If Instr(satac_code, ",") > 0 Then
						Dim splitmystring : splitmystring = split(satac_code,",")
						satac_code1 = splitmystring(0)
						satac_code2 = splitmystring(1)
						sql = "SELECT * from ATAR WHERE SATAC_CODE = '" & satac_code1 & "' ORDER BY SRT"
						sql2 = "SELECT * from ATAR WHERE SATAC_CODE = '" & satac_code2 & "' ORDER BY SRT"
					Else
						sql = "SELECT * from ATAR WHERE SATAC_CODE = '" & satac_code & "' ORDER BY SRT"
					End If

					rs.Open sql, conn

		If not (rs.EOF) Then

		%>
			<div class="Pagedivider">&nbsp;</div>

			<div class="logo">
				<img src="uni-logo.jpg" class="heading">
				<h1 class="heading">ATAR Profile</h1>
			</div>


		<%
		If Instr(satac_code, ",") > 0 Then
			Response.write("<h3 class=""programName heading"">PROGRAM NAME : <span>" & rs.Fields("ProgramName") & "</span>&nbsp;/&nbsp;Program Code : <span>" & progCode & "</span>&nbsp;/&nbsp;SATAC : <span>" & satac_code1 & "</span></h3>")
		Else
				Response.write("<h3 class=""programName heading"">PROGRAM NAME : <span>" & rs.Fields("ProgramName") & "</span>&nbsp;/&nbsp;Program Code : <span>" & progCode & "</span>&nbsp;/&nbsp;SATAC : <span>" & satac_code & "</span></h3>")
		End If
		%>

		<blockquote>
			The table below relates to all applicants whose admission is based mostly on secondary education undertaken within the previous two years, and who were selected on the basis of their ATAR alone. The table includes ATAR (excluding any adjustment factors) and Selection Rank (ATAR plus any adjustment factors).
				<ul><b>Note:</b>
					<li>*L/N - indicates low numbers if less than 5 ATAR-based offers made</li>
					<li>#N/P - indicates figure is not published if less than 25 ATAR-based offers made</li>
				</ul>
		</blockquote>

		<table class="pure-table  pure-table-bordered equalDivide" width="100%">
			<thead>
			<tr>
				<th>
					<p>(ATAR-based offers only, across all offer rounds)</p>
				</th>
				<th>
					<p>ATAR (OP in QLD)<span style="font-size:smaller;">&nbsp; (Excluding adjustment factors)</span></p>
				</th>
				<th>
					<p>Selection Rank<span style="font-size:smaller;">&nbsp;(ATAR/OP plus any adjustment factors)</span></p>
				</th>
			</tr>
		</thead>

				<% do until rs.EOF

					finalGroupingA = rs.Fields("DESCRIPTION")
					if Left(finalGroupingA, 4) = "High" Then
						Response.Write("<tr><td>" & rs.Fields("DESCRIPTION")) & " *</td>"
						Response.Write("<td>" & rs.Fields("ATAR")) & "</td>"
						Response.Write("<td>" & rs.Fields("SelectionRank")) & "</td></tr>"
					End If

					finalGroupingB = rs.Fields("DESCRIPTION")
					if Left(finalGroupingB, 4) = "75th" Then
						Response.Write("<tr><td>" & rs.Fields("DESCRIPTION")) & " #</td>"
						Response.Write("<td>" & rs.Fields("ATAR")) & "</td>"
						Response.Write("<td>" & rs.Fields("SelectionRank")) & "</td></tr>"
					End If

					finalGroupingC = rs.Fields("DESCRIPTION")
					if Left(finalGroupingC, 4) = "Medi" Then
						Response.Write("<tr><td>" & rs.Fields("DESCRIPTION")) & " #</td>"
						Response.Write("<td>" & rs.Fields("ATAR")) & "</td>"
						Response.Write("<td>" & rs.Fields("SelectionRank")) & "</td></tr>"
					End If

					finalGroupingD = rs.Fields("DESCRIPTION")
					if Left(finalGroupingD, 4) = "25th" Then
						Response.Write("<tr><td>" & rs.Fields("DESCRIPTION")) & " #</td>"
						Response.Write("<td>" & rs.Fields("ATAR")) & "</td>"
						Response.Write("<td>" & rs.Fields("SelectionRank")) & "</td></tr>"
					End If

					finalGroupingE = rs.Fields("DESCRIPTION")
					if Left(finalGroupingE, 4) = "Lowe" Then
						Response.Write("<tr><td>" & rs.Fields("DESCRIPTION")) & " *</td>"
						Response.Write("<td>" & rs.Fields("ATAR")) & "</td>"
						Response.Write("<td>" & rs.Fields("SelectionRank")) & "</td></tr>"
					End If

				rs.MoveNext
				loop %>
			</table>

			<%
					If Instr(satac_code, ",") > 0 Then
					set rs = Server.CreateObject("ADODB.recordset")
					sql = "SELECT * from ATAR WHERE SATAC_CODE = '" & satac_code2 & "' ORDER BY SRT"

					rs.Open sql, conn
					If Not rs.eof Then
						Response.write("<br><h3 class=""programName heading"">PROGRAM NAME : <span>" & rs.Fields("ProgramName") & "</span>&nbsp;/&nbsp;Program Code : <span>" & progCode & "</span>&nbsp;/&nbsp;SATAC : <span>" & satac_code2 & "</span></h3><br>")
			%>

			<table class="pure-table  pure-table-bordered equalDivide" width="100%">
				<thead>
				<tr>
					<th>
						<p>(ATAR-based offers only, across all offer rounds)</p>
					</th>
					<th>
						<p>ATAR (OP in QLD)<span style="font-size:smaller;">&nbsp; (Excluding adjustment factors)</span></p>
					</th>
					<th>
						<p>Selection Rank<span style="font-size:smaller;">&nbsp;(ATAR/OP plus any adjustment factors)</span></p>
					</th>
				</tr>
			</thead>

				<% do until rs.EOF

					finalGroupingA = rs.Fields("DESCRIPTION")
					if Left(finalGroupingA, 4) = "High" Then
						Response.Write("<tr><td>" & rs.Fields("DESCRIPTION")) & " *</td>"
						Response.Write("<td>" & rs.Fields("ATAR")) & "</td>"
						Response.Write("<td>" & rs.Fields("SelectionRank")) & "</td></tr>"
					End If

					finalGroupingB = rs.Fields("DESCRIPTION")
					if Left(finalGroupingB, 4) = "75th" Then
						Response.Write("<tr><td>" & rs.Fields("DESCRIPTION")) & " #</td>"
						Response.Write("<td>" & rs.Fields("ATAR")) & "</td>"
						Response.Write("<td>" & rs.Fields("SelectionRank")) & "</td></tr>"
					End If

					finalGroupingC = rs.Fields("DESCRIPTION")
					if Left(finalGroupingC, 4) = "Medi" Then
						Response.Write("<tr><td>" & rs.Fields("DESCRIPTION")) & " #</td>"
						Response.Write("<td>" & rs.Fields("ATAR")) & "</td>"
						Response.Write("<td>" & rs.Fields("SelectionRank")) & "</td></tr>"
					End If

					finalGroupingD = rs.Fields("DESCRIPTION")
					if Left(finalGroupingD, 4) = "25th" Then
						Response.Write("<tr><td>" & rs.Fields("DESCRIPTION")) & " #</td>"
						Response.Write("<td>" & rs.Fields("ATAR")) & "</td>"
						Response.Write("<td>" & rs.Fields("SelectionRank")) & "</td></tr>"
					End If

					finalGroupingE = rs.Fields("DESCRIPTION")
					if Left(finalGroupingE, 4) = "Lowe" Then
						Response.Write("<tr><td>" & rs.Fields("DESCRIPTION")) & " *</td>"
						Response.Write("<td>" & rs.Fields("ATAR")) & "</td>"
						Response.Write("<td>" & rs.Fields("SelectionRank")) & "</td></tr>"
					End If

				rs.MoveNext
				loop %>
			</table>
			<%
					End If
				End If
			End If
			%>

		</body>
</html>
