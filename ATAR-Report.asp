<!DOCTYPE html>
<html>

<head>
	<title>Test website - Note content will be regularly deleted from this site with no warning.</title>
	<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	<meta name="keywords" content="test website, web testing">
	<meta name="description" content="Test website - Note content will be regularly deleted from this site with no warning.">
	<meta name="robots" content="all">
	<link rel="stylesheet" href="https://unpkg.com/purecss@1.0.0/build/pure-min.css" integrity="sha384-nn4HPE8lTHyVtfCBi5yW9d20FjT8BJwUXyWZT9InLYax14RDjBj46LmSztkmNP9w" crossorigin="anonymous">
	<link rel="stylesheet" type="text/css" href="transparency.css">
</head>

<body>

	<h1>Transparency Project</h1>
	<h2>ATAR Report</h2>
			
			<%
					Dim conn
					Set conn = Server.CreateObject("ADODB.Connection")
					conn.Open "Provider=SQLOLEDB;Data Source=.\SQLEXPRESS;Initial Catalog=transparency;User Id=sa;Password=KLF-chill-out-1990"
					If conn.errors.count = 0 Then
						Response.Write "<h3 class=""success"">Database Connected OK</h3>"
					End If
					
					satac = Request.QueryString("satac")
					response.write("<strong>SATAC Query= " & satac & "</strong>")
			
					set rs = Server.CreateObject("ADODB.recordset")
					sql = "SELECT * from ATAR WHERE SATAC_CODE = '" & satac & "' ORDER BY SRT"
					rs.Open sql, conn
					
			%>
		<!-- check we have data-->
		
		
		<% 
		If (rs.EOF) Then 
			Response.Write("<h4 class=""error"">Errors - Query returned no results for SATAC Code : ''" & satac & "''</h4>")
		Else
		
		Response.write("<h3 class=""programName"">" & rs.Fields("ProgramName") & "</h3>") %>			
			
			<table class="pure-table  pure-table-bordered equalDivide" width="80%">
				<thead>
				<tr>
			    <th>
						<p>(ATAR-based offers only, across all offer rounds)</p>
						<p style="font-size:smaller;">[Note: this table relates to all students selected on the basis of ATAR alone or ATAR in combination with other factors. To ensure comparability across all providers, the "ATAR" figures used must reflect the original unadjusted figures without the impact of 'bonus points' or other adjustments. "Selection Rank" figures (if used) will reflect the same cohort but including the impact of 'bonus points' or other adjustments.</p>
						<p style="font-size:smaller;">Students selected on the basis of special consideration or otherwise not on the basis of their ATAR should not be included in this table.]</p>
					</th>
			    <th>
						<p>ATAR (OP in QLD)</p><p style="font-size:smaller;">(Excluding adjustment factors)</p><p style="font-size:smaller;">[required]</p>
						<p style="font-size:smaller;">[NB: Raw ATAR profile for all students offered a place wholly or partly on the basis of ATAR]	</p>
					</th>
			    <th>
						<p>Selection Rank</p>
						<p style="font-size:smaller;">(ATAR/OP plus any adjustment factors)</p>
						<p style="font-size:smaller;">[optional / only if relevant]</p>
						<p style="font-size:smaller;">[NB: Selection Rank profile for the same students as in previous column]</p>
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
			%>
	    
		</body>
</html>