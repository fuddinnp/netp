<%@ LANGUAGE="VBScript" %>
<%	'**************************************************************************
	'* ASP FormMail                                                           *
	'*                                                                        *
	'* Do not remove this notice.                                             *
	'*                                                                        *
	'* Copyright 1999-2008 by Mike Hall.                                      *
	'* Please see http://www.brainjar.com for documentation and terms of use. *
	'**************************************************************************

	'- Customization of these values is required, see documentation. ----------

	mailComp   = "CDOSYS"
	'mailComp   = "CDOSYS"
	smtpServer = "mail.netpace.com"
	fromAddr   = "support@netpace.com"

	'allowedHosts      = Array("www.example.net", "example.net")
	allowedHosts      = Array()
	allowedRecipients = Array()
	allowedEnvars     = Array("HTTP_USER_AGENT", "REMOTE_ADDR", "REMOTE_USER")

	allowCcToFlag = true

	botCheckFlag    = false
	botCheckID      = "MyBotCheckID"
	botCheckMinTime = 5

	'- End required customization section. ------------------------------------

	'Initialize.
	Response.Buffer = true
	errorMsgs = Array()

	'Check for form data.
	if Request.ServerVariables("Content_Length") = 0 then
		call AddErrorMsg("No form data submitted.")
	end if

	'If bot checking is enabled, check the elapsed time.
	if botCheckFlag then
		startTime = Session(botCheckID)
		if not IsDate(startTime) then
			call AddErrorMsg("Invalid submission.")
		elseif DateDiff("s", startTime, Now()) < botCheckMinTime then
			call AddErrorMsg("Invalid submission.")
		end if
	end if

	'Check if the referering host is allowed.
	if UBound(allowedHosts) >= 0 then
		host = GetHost(Request.ServerVariables("HTTP_REFERER"))
		if host = "" then
			call AddErrorMsg("No referer.")
		elseif not InList(host, allowedHosts) then
			call AddErrorMsg("Unauthorized host: '" & host & "'.")
		end if
	end if

	if Request.Form("_g-recaptcha-response") = "" then
		call AddErrorMsg("Captcha is empty.")
	else
		dim post_data, objHTTP
		dim recaptcha_response
		recaptcha_response = Request.Form("_g-recaptcha-response")
		post_data = "secret=" & "6LccLNcjAAAAAFMRQO9a7_0l7Ml9nyAnXRQXdqbK" & "&response=" & recaptcha_response & "&remoteip=" & Request.ServerVariables("REMOTE_ADDR")

		set objHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objHTTP.open "POST", "https://www.google.com/recaptcha/api/siteverify", false
		objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objHTTP.send post_data

		if objHTTP.status = 200 then
		    dim response_text
		    response_text = objHTTP.responseText
		    if InStr(response_text, "true") > 0 then
		        ' reCAPTCHA was verified successfully
		    else
		        call AddErrorMsg("Captcha not validated.")
		    end if
		else
		    call AddErrorMsg("Captcha not validated.")
		end if
	end if

	'Check the email recipients.
	if Request.Form("_recipients") = "" then
		call AddErrorMsg("No email recipient(s) specified.")
	else
		recipients = Split(Request.Form("_recipients"), ",")
		for each addr in recipients
			addr = Trim(addr)
			if not IsValidEmailAddress(addr) then
				call AddErrorMsg("Invalid email address in recipient list: " & addr & ".")
			elseif UBound(allowedRecipients) >= 0 then
				if not inList(addr, allowedRecipients) then
					call AddErrorMsg("Unauthorized email address in recipient list: " & addr & ".")
				end if
			end if
		next
		recipients = Join(recipients, ",")
	end if
		
	if Request.Form("Email Address") = "" then
		call AddErrorMsg("No email recipient(s) specified.")
	else
		toAddress1 = Split(Request.Form("Email Address"), ",")
		for each addr in toAddress1
			addr = Trim(addr)
			if not IsValidEmailAddress(addr) then
				call AddErrorMsg("Invalid email address in recipient list: " & addr & ".")
			elseif UBound(allowedRecipients) >= 0 then
				if not inList(addr, allowedRecipients) then
					call AddErrorMsg("Unauthorized email address in recipient list: " & addr & ".")
				end if
			end if
		next
		toAddress1 = Join(toAddress1, ",")
	end if

	'Check for a cc-to or reply-to address.
	ccToAddr = ""
	replyToAddr = ""
	name = Trim(Request.Form("_ccToField"))
	if name <> "" then
		if not allowCcToFlag then
			call AddErrorMsg("Configuration error: use of _ccToField not permitted.")
		else
			ccToAddr = Request.Form(name)
			if ccToAddr <> "" then
				if not IsValidEmailAddress(ccToAddr) then
					call AddErrorMsg("Invalid email address in " & name & " field: " & ccToAddr & ".")
				end if
			end if
		end if
	else
		name = Trim(Request.Form("_replyToField"))
		if name <> "" then
			replyToAddr = Request.Form(name)
			if replyToAddr <> "" then
				if not IsValidEmailAddress(replyToAddr) then
					call AddErrorMsg("Invalid email address in " & name & " field: " & replyToAddr & ".")
				end if
			end if
		end if
	end if

	'Get the subject text.
	subject = Request.Form("_subject")

	'If required fields are specified, check them.
	if Request.Form("_requiredFields") <> "" then
		required = Split(Request.Form("_requiredFields"), ",")
		for each name in required
			name = Trim(name)
			if Left(name, 1) <> "_" and Request.Form(name) = "" then
				call AddErrorMsg("Missing value for " & name)
			end if
		next
	end if

	'If a field order was given, use it. Otherwise use the order the fields
 	'were received in.
	str = ""
	if Request.Form("_fieldOrder") <> "" then
		fieldOrder = Split(Request.Form("_fieldOrder"), ",")
		for each name in fieldOrder
			if str <> "" then
				str = str & ","
			end if
			str = str & Trim(name)
		next
		fieldOrder = Split(str, ",")
	else
		fieldOrder = FormFieldList()
	end if

	'If there were no errors, build the email note and send it.
	if UBound(errorMsgs) < 0 then

		'Build table of form fields and values.
		body = "<table border=""0"" cellpadding=""2"" cellspacing=""0"">" & vbCrLf
		for each name in fieldOrder
			body = body _
			     & "<tr valign=""top"">" _
			     & "<td><strong>" & name & ":</strong></td>" _
			     & "<td>" & Request.Form(name) & "</td>" _
			     & "</tr>" & vbCrLf
		next
		body = body & "</table>" & vbCrLf

		'Add a table for any requested environment variables.
		if Request.Form("_envars") <> "" then
			body = body _
			     & "<p>&nbsp;</p>" & vbCrLf _
			     & "<table border=""0"" cellpadding=""2"" cellspacing=""0"">" & vbCrLf
			envars = Split(Request.Form("_envars"), ",")
			for each name in envars
				name = Trim(name)

				'Only show environment variables in the permitted list.
				showEnvar = true
				if UBound(allowedEnvars) >= 0 then
					showEnvar = InList(name, allowedEnvars)
				end if
				if showEnvar then
					body = body _
					     & "<tr valign=""top"">" _
					     & "<td><strong>" & name & ":</strong></td>" _
					     & "<td>" & Request.ServerVariables(name) & "</td>" _
					     & "</tr>" & vbCrLf
				end if
			next
			body = body & "</table>" & vbCrLf
		end if

		'Send it.
		str = SendMail()
		if str <> "" then
			AddErrorMsg(str)
		else
			'Clear the bot check timestamp.
			Session.Contents.Remove(botCheckID)

			'Redirect if a URL was given.
			if Request.Form("_redirectUrl") <> "" then
				Response.Redirect(Request.Form("_redirectUrl"))
			end if
		end if

	end if %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head><link rel="SHORTCUT ICON" href="images/netpace-icon.ico" />
<link rel="icon" type="image/ico" href="images/netpace-icon.ico" />
	<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
	<title>Form Mail</title>
	<style type="text/css">
		body
		{
			background-color: #ffffff;
			color: #000000;
			font-family: Arial, Helvetica, sans-serif;
			font-size: 10pt;
		}
		
		table
		{
			border: solid 1px #000000;
			border-collapse: collapse;
		}
		
		td, th
		{
			border: solid 1px #000000;
			font-family: Arial, Helvetica, sans-serif;
			font-size: 10pt;
			padding: 2px 8px;
		}
		
		th
		{
			background-color: #c0c0c0;
		}
		
		.error
		{
			color: #c00000;
		}
	</style>
<script type="text/javascript">

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-29276894-1']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>
</head>
<body>
<%	if UBound(errorMsgs) >= 0 then %>
	<p class="error">
	Form could not be processed due to the following errors:</p>
	<ul>
<%		for each msg in errorMsgs %>
		<li class="error"><% = msg %></li>
<%		next %>
	</ul>
	<p><a href="#" onclick="history.go(-1); return false;">Back</a></p>
<%	else %>
	<table cellpadding="0" cellspacing="0">
		<tr>
			<th colspan="2" valign="bottom">Thank you, the following information has been sent:</th>
		</tr>
<%		for each name in fieldOrder %>
		<tr valign="top">
			<td><strong><% = name %></strong></td>
			<td><% = Request.Form(name) %></td>
		</tr>
<%		next %>
	</table>
<%		if Request.Form("_continueUrl") <> "" then %>
	<p><a href="<% = Request.Form("_continueUrl") %>">Continue</a></p>
<%		end if
	end if %>
</body>
</html>
<%	'==========================================================================
	' Functions and subroutines.
	'==========================================================================

	'--------------------------------------------------------------------------
	' Adds an error message to the list.
	'--------------------------------------------------------------------------
	sub AddErrorMsg(msg)

		dim n

		n = UBound(errorMsgs)
		Redim Preserve errorMsgs(n + 1)
		errorMsgs(n + 1) = msg

	end sub

	'--------------------------------------------------------------------------
	' Extracts the host name from a URL.
	'--------------------------------------------------------------------------
	function GetHost(url)

		dim i, str

		GetHost = ""

		'Strip down to host or IP address and port number, if any.
		if Left(url, 7) = "http://" then
			str = Mid(url, 8)
		elseif Left(url, 8) = "https://" then
			str = Mid(url, 9)
		end if
		i = InStr(str, "/")
		if i > 1 then
			str = Mid(str, 1, i - 1)
		end if
		GetHost = str

	end function

	'--------------------------------------------------------------------------
	' Returns true if the given string is in the given array.
	'--------------------------------------------------------------------------
	function InList(str, list)

		dim item

		InList = false

		'Scan the list.
		for each item in list
			if str = item then
				InList = true
				exit function
			end if
		next

	end function

	'--------------------------------------------------------------------------
	' Returns true if the given email address is in valid format.
	'--------------------------------------------------------------------------
	function IsValidEmailAddress(addr)

		dim list, item
		dim i, c

		IsValidEmailAddress = true

		'Exclude any address with '..'.
		if InStr(addr, "..") > 0 then
			IsValidEmailAddress = false
			exit function
		end if

		'Split email address into the user and domain names.
		list = Split(addr, "@")
		if UBound(list) <> 1 then
			IsValidEmailAddress = false
			exit function
		end if

		'Check both names.
		for each item in list

			'Make sure the name is not zero length.
			if Len(item) <=  0 then
				IsValidEmailAddress = false
				exit function
			end if

			'Make sure only valid characters appear in the name.
			for i = 1 to Len(item)
				c = Lcase(Mid(item, i, 1))
				if InStr("abcdefghijklmnopqrstuvwxyz&_-.", c) <= 0 and not IsNumeric(c) then
					IsValidEmailAddress = false
					exit function
				end if
			next

			'Make sure the name does not start or end with invalid characters.
			if Left(item, 1) = "." or Right(item, 1) = "." then
				IsValidEmailAddress = false
				exit function
			end if

		next

		'Check for a '.' character in the domain name.
		if InStr(list(1), ".") <= 0 then
			IsValidEmailAddress = false
			exit function
		end if

	end function

	'--------------------------------------------------------------------------
	' Builds an array of form field names ordered as they were received.
	' Note that fields whose name starts with an underscore are ignored.
	'--------------------------------------------------------------------------
	function FormFieldList()

		dim str, i, name

		str = ""
		for i = 1 to Request.Form.Count
			for each name in Request.Form
				if Left(name, 1) <> "_" and Request.Form(name) is Request.Form(i) then
					if str <> "" then
						str = str & ","
					end if
					str = str & name
					exit for
				end if
			next
		next
		FormFieldList = Split(str, ",")

	end function

	'--------------------------------------------------------------------------
	' Sends email based on mail component. Uses global variables for parameters
	' because there are so many.
	'--------------------------------------------------------------------------

	function SendMail()

		dim mailObj, cdoMessage, cdoConfig
		dim addrList

		SendMail = ""

		'Send email (CDONTS version). Note: CDONTS has no error checking.
		if mailComp = "CDONTS" then
			set mailObj = Server.CreateObject("CDONTS.NewMail")
			mailObj.BodyFormat = 0
			mailObj.MailFormat = 0
			mailObj.From = fromAddr
			mailObj.To = recipients
			if ccToAddr <> "" then
				mailObj.Value("Reply-To") = Trim(ccToAddr)
				mailObj.CC = Trim(ccToAddr)
			elseif replyToAddr <> "" then
				mailObj.Value("Reply-To") = Trim(replyToAddr)
			end if
			mailObj.Subject = subject
			mailObj.Body = body
			mailObj.Send
			set mailObj = Nothing
			exit function
		end if

		'Send email (CDOSYS version).
		if mailComp = "CDOSYS" then
			set cdoMessage = Server.CreateObject("CDO.Message")
			set cdoConfig = Server.CreateObject("CDO.Configuration")
			cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer
  			'*** Authentication ***'  
			cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1  
			
			
			
			'************ Replace your email and password  here **********************'
			
			
			cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "inforequest@netpace.com" 	'*** email address ***'  
			cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "P@k1stan15inSOUTHAsi@" 	'*** Password ***' 
			
			
			'**************************************************************************'


			
			cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30  

			cdoConfig.Fields.Update
			set cdoMessage.Configuration = cdoConfig
			cdoMessage.From =  fromAddr
			cdoMessage.To = recipients
			if ccToAddr <> "" then
				cdoMessage.ReplyTo = Trim(ccToAddr)
				cdoMessage.CC = Trim(ccToAddr)
			elseif replyToAddr <> "" then
				cdoMessage.ReplyTo = Trim(replyToAddr)
			end if
			cdoMessage.Subject = subject
			cdoMessage.HtmlBody = body
			on error resume next
			cdoMessage.Send
			if Err.Number <> 0 then
				SendMail = "Email send failed: " & Err.Description & "."
			end if
			
			cdoMessage.To = toAddress1
			if ccToAddr <> "" then
				cdoMessage.ReplyTo = Trim(ccToAddr)
				cdoMessage.CC = Trim(ccToAddr)
			elseif replyToAddr <> "" then
				cdoMessage.ReplyTo = Trim(replyToAddr)
			end if
			cdoMessage.Subject = "We've received your email."
			cdoMessage.HtmlBody = "Hi,<br/><br/>Thank you for your interest in Netpace.<br/>So this is to let you know, we have received your message and we'll respond back shortly. <br/><br/>Sincerely,<br/>Netpace Staff"
			on error resume next
			cdoMessage.Send
			if Err.Number <> 0 then
				SendMail = "Email send failed: " & Err.Description & "."
			end if
			
			set cdoMessage = Nothing
			set cdoConfig = Nothing
			exit function
		end if
		
		'Send email (JMail version).
		if mailComp = "JMail" then
			set mailObj = Server.CreateObject("JMail.SMTPMail")
			mailObj.Silent = true
			mailObj.ServerAddress = smtpServer
			mailObj.Sender = fromAddr
			mailObj.Subject = subject
			addrList = Split(recipients, ",")
			for each addr in addrList
				mailObj.AddRecipient Trim(addr)
			next
			if ccToAddr <> "" then
				mailObj.ReplyTo = Trim(ccToAddr)
				mailObj.AddRecipientCC Trim(ccToAddr)
			elseif replyToAddr <> "" then
				mailObj.ReplyTo = Trim(replyToAddr)
			end if
			mailObj.ContentType = "text/html"
			mailObj.Body = body
			if not mailObj.Execute then
				SendMail = "Email send failed: " & mailObj.ErrorMessage & "."
			end if
			exit function
		end if

		'Send email (ASPMail version).
		if mailComp = "ASPMail" then
			set mailObj = Server.CreateObject("SMTPsvg.Mailer")
			mailObj.RemoteHost  = smtpServer
			mailObj.FromAddress = fromAddr
			for each addr in Split(recipients, ",")
				mailObj.AddRecipient "", Trim(addr)
			next
			if ccToAddr <> "" then
				mailObj.ReplyTo = Trim(ccToAddr)
				mailObj.AddCC "", Trim(ccToAddr)
			elseif replyToAddr <> "" then
				mailObj.ReplyTo = Trim(replyToAddr)
			end if
			mailObj.Subject = subject
			mailObj.ContentType = "text/html"
			mailObj.BodyText = body
			if not mailObj.SendMail then
				SendMail = "Email send failed: " & mailObj.Response & "."
			end if
			exit function
		end if

	end function %>

