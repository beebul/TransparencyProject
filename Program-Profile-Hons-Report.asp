<!DOCTYPE html>
<html>
<div class="logo">
    <img src="uni-logo.jpg" class="heading">
    <h1 class="heading">ADMISSIONS INFORMATION SET</h1>
</div>
<%
			Dim conn
			Set conn = Server.CreateObject("ADODB.Connection")
			conn.Open "Provider=SQLOLEDB;Data Source=.\SQLEXPRESS;Initial Catalog=transparency;User Id=sa;Password=KLF-chill-out-1990"

			ppc = Request.QueryString("ppc")

			set rs = Server.CreateObject("ADODB.recordset")
			sql = " SELECT * from ProgramProfileHons WHERE ProgramPlanCode = '" & ppc & "' ORDER BY FINAL_GROUPING"
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

	Response.write("<h3 class=""programName heading""> Program Name : <span>" & rs.Fields("ProgramName") & "</span>&nbsp;/&nbsp;Program Code : <span>" & progCode & "</span>")
	If satac_code <> "" Then
		Response.Write(" / SATAC CODE: " & satac_code)
	End If

	Response.Write("</h3>	<br>")

	If (rs.Fields("Offering") <> "") Then
		offeringType = Trim(rs.Fields("Offering"))
		fullOfferingType= expandOT(offeringType)
		Response.Write("<h3 class=""programName heading"" style=""margin-top: -1em;color:#01029A;"">" & fullOfferingType & "</h3>")
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
        <blockquote>The admissions information set for each degree outlines a collection of information that will help you to gauge and compare study options, course admission criteria and your likely student peer cohort across multiple courses and admission options. Each admissions information set at the degree level will include a student profile for the degree and, for most degrees, an ATAR profile. An ATAR profile will not be available where there were no applicants in the secondary education applicant group who were admitted solely on the basis of their ATAR for that degree.
            <br>
            <br><b>Student profile</b>
            <br> The table below gives an indication of the likely peer cohort for new students at UniSA. It provides data of students who commenced undergraduate study in the most recent intake period.
            <ul>
                <b>Note:</b>
                <li>L/N - low numbers: the numbers of students is less than 5</li>
                <li>N/A - data not available for this item</li>
                <li>N/P – Not published: the number is hidden to prevent calculation of numbers in cells with less than 5 students.</li>
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
		%>
    </body>

</html>