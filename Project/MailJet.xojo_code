#tag Class
Protected Class Mailjet
	#tag Method, Flags = &h21
		Private Function ConvertAttachments(oEmail as EmailMessage) As JSONItem
		  var jsAttachmentsArray as new JSONItem("[]")
		  
		  for each oAttachment as EmailAttachment in oEmail.Attachments
		    var jsAttachment as new JSONItem
		    jsAttachment.Value("ContentType") = oAttachment.MIMEType
		    jsAttachment.Value("Filename") = oAttachment.Name
		    jsAttachment.Value("Base64Content") = oAttachment.Data
		    
		    jsAttachmentsArray.Add(jsAttachment)
		    
		  next oAttachment
		  
		  return jsAttachmentsArray
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ConvertEmailToMailjet(oMail as EmailMessage) As JSONItem
		  // Validate From
		  if oMail.FromAddress.Trim = "" then
		    var ex as new MailjetException
		    ex.Message = "EmailMessage has no FromAddress"
		    RaiseEvent Error(ex)
		    return nil
		    
		  end
		  
		  // Validate To
		  if oMail.ToAddress.Trim = "" then
		    var ex as new MailjetException
		    ex.Message = "EmailMessage has no Recepients"
		    RaiseEvent Error(ex)
		    return nil
		    
		  end
		  
		  var jsFromArray as JSONItem = GetAddressArray(oMail.FromAddress)
		  
		  var jsToArray as JSONItem = GetAddressArray(oMail.ToAddress)
		  var jsCCArray as JSONItem = GetAddressArray(oMail.CCAddress)
		  var jsBCCArray as JSONItem = GetAddressArray(oMail.BCCAddress)
		  
		  // Start building JSON item
		  var jsBody as new JSONItem
		  jsBody.Value("From") = jsFromArray(0)
		  jsBody.Value("To") = jsToArray
		  
		  if jsCCArray.Count > 0 then
		    jsBody.Value("CC") = jsCCArray
		    
		  end
		  
		  if jsBCCArray.Count > 0 then
		    jsBody.Value("BCC") = jsBCCArray
		    
		  end
		  
		  jsBody.Value("Subject") = oMail.Subject
		  
		  if oMail.BodyPlainText <> "" then
		    jsBody.Value("TextPart") = oMail.BodyPlainText
		    
		  end
		  
		  if oMail.BodyHTML <> "" then
		    jsBody.Value("HTMLPart") = oMail.BodyHTML
		    
		  end
		  
		  // Attachments are not handled at this time
		  if oMail.Attachments.LastIndex > -1 then
		    jsBody.Value("Attachments") = ConvertAttachments(oMail)
		    
		  end
		  
		  return jsBody
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetAddressArray(sRecipientsCSV as String) As JSONItem
		  var jsToArray as new JSONItem("[]")
		  
		  var arsAddressed() as String = sRecipientsCSV.Split(",")
		  
		  for each sAddress as String in arsAddressed
		    var jsToItem as new JSONItem
		    
		    var rx as new RegEx
		    rx.SearchPattern = kRxEmail
		    
		    var rxm as RegExMatch = rx.Search(sAddress)
		    if rxm <> nil then
		      jsToItem.Value("Email") = rxm.SubExpressionString(1)
		      
		    end
		    
		    // Now check for name
		    rx = new RegEx
		    rx.SearchPattern = kRxEmailName
		    rxm = rx.Search(sAddress)
		    
		    if rxm <> nil then
		      jsToItem.Value("Name") = rxm.SubExpressionString(1)
		      
		    end
		    
		    // Add it to the array
		    jsToArray.Add(jsToItem)
		    
		  next sAddress
		  
		  return jsToArray
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetFromObject(oMail as EmailMessage) As Dictionary
		  // Set up From object
		  var dictFrom as new Dictionary
		  
		  var rx as new RegEx
		  rx.SearchPattern = kRxEmail
		  
		  var rxm as RegExMatch = rx.Search(oMail.FromAddress)
		  if rxm <> nil then
		    dictFrom.Value("Email") = rxm.SubExpressionString(1)
		    
		  end
		  
		  // Now check for name
		  rx = new RegEx
		  rx.SearchPattern = kRxEmailName
		  rxm = rx.Search(oMail.FromAddress)
		  
		  if rxm <> nil then
		    dictFrom.Value("Name") = rxm.SubExpressionString(1)
		    
		  end
		  
		  return dictFrom
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleError(oSender as URLConnection, ex as RuntimeException)
		  #pragma unused oSender
		  
		  // Raise Error event.
		  Error(ex)
		  
		  mbBusy = false
		  moSock = nil
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleResponse(oSender As URLConnection, URL As String, HTTPStatus As Integer, content As String)
		  #pragma unused oSender
		  #pragma unused URL
		  
		  if HTTPStatus <> 200 then
		    var ex as new MailjetException
		    ex.Message = "HTTP response code was not okay: " + HTTPStatus.ToString
		    
		    try
		      // Attempt to parse out the error messager
		      var vResponse as Variant = ParseJSON(content.DefineEncoding(Encodings.UTF8))
		      var dictResponse as Dictionary = Dictionary(vResponse)
		      
		      if dictResponse.HasKey("ErrorMessage") then
		        ex.Message = dictResponse.Value("ErrorMessage")
		        
		      end
		      
		    catch ex2 as IllegalCastException
		      // Trying to turn the variant into a dictionary failed
		      // This can happen if it's not the single json object we're expecting
		      
		    catch ex2 as InvalidJSONException
		      // Response wasn't json
		      // Not much we can parse here
		      
		    end try
		    
		    RaiseEvent Error(ex)
		    
		  else
		    // Good response, pass to parser who will parse out
		    // individual errors and then raise MailSent
		    HandleResponse200(content.DefineEncoding(Encodings.UTF8))
		    
		  end
		  
		  mbBusy = false
		  moSock = nil
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleResponse200(sBody as String)
		  var dictResponse, ardictMessagesResponse() as Dictionary
		  
		  try
		    // Convert json body to Xojo object
		    var vBody as Variant = ParseJSON(sBody)
		    dictResponse = Dictionary(vBody)
		    
		    if dictResponse.HasKey("Messages") then
		      var arvMessages() as Variant = dictResponse.Value("Messages")
		      for each vMessage as Variant in arvMessages
		        ardictMessagesResponse.Add(Dictionary(vMessage))
		        
		      next vMessage
		      
		    end
		    
		  catch ex as IllegalCastException
		    // Trying to turn the variant into a dictionary failed
		    // This can happen if it's not the single json object we're expecting
		    RaiseEvent Error(ex)
		    return
		    
		  catch ex as InvalidJSONException
		    // Response wasn't json
		    // Not much we can parse here
		    RaiseEvent Error(ex)
		    return
		    
		  end try
		  
		  // Store a flag for if the response is all failures
		  // That shouldn't raise a MailSent event because no mail was ever sent
		  var bAllFailures as Boolean = True
		  
		  for each dictMessageResult as Dictionary in ardictMessagesResponse
		    var sStatus as String =dictMessageResult.Lookup("Status", "unknown")
		    
		    select case sStatus
		    case "success"
		      bAllFailures = false
		      
		    case else
		      var ex as new MailjetException
		      ex.Message = "A message failed to send: " + GenerateJSON(dictMessageResult)
		      Error(ex)
		      
		    end
		    
		  next dictMessageResult
		  
		  // Raise Sent event if something sent!
		  if not bAllFailures then
		    RaiseEvent MailSent
		    
		  end
		  
		  // Cleanup
		  RemoveAllMessages
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function NewSocket() As URLConnection
		  // Create URLConnection and add the authentication header
		  var sAuth as String = EncodeBase64(kAPIKey + ":" + kAPISecret, 0)
		  
		  var oSock as new URLConnection
		  oSock.RequestHeader("Authorization") = "Basic " + sAuth
		  
		  // Handle server responses that aren't 200
		  AddHandler oSock.Error, WeakAddressOf HandleError
		  AddHandler oSock.ContentReceived, WeakAddressOf HandleResponse
		  
		  return oSock
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub RemoveAllMessages()
		  me.Messages.ResizeTo(-1)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SendMail()
		  // Validate
		  if mbBusy then
		    var ex as new MailjetException
		    ex.Message = "This Mailjet socket is already in use, please wait for the MailSent event."
		    RaiseEvent Error(ex)
		    return
		    
		  end
		  
		  var jsRequest as new JSONItem
		  
		  #if DebugBuild then
		    // This SandboxMode flag prevents emails from actually sending
		    // Take this out if you're testing actual delivery
		    jsRequest.Value("SandboxMode") = true
		    
		  #endif
		  
		  var jsMessages as new JSONItem("[]")
		  
		  // Put all messages into request
		  for each oEmail as EmailMessage in Messages
		    var jsMail as JSONItem = ConvertEmailToMailjet(oEmail)
		    jsMessages.Add(jsMail)
		    
		  next oEmail
		  
		  jsRequest.Value("Messages") = jsMessages
		  
		  var sBody as String = jsRequest.ToString
		  
		  mbBusy = true
		  moSock = NewSocket
		  
		  moSock.SetRequestContent(sBody, "application/json")
		  
		  moSock.Send("POST", "https://api.mailjet.com/v3.1/send")
		End Sub
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event Error(ex as RuntimeException)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MailSent()
	#tag EndHook


	#tag Property, Flags = &h21
		Private mbBusy As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		Messages() As EmailMessage
	#tag EndProperty

	#tag Property, Flags = &h21
		Private moSock As URLConnection
	#tag EndProperty


	#tag Constant, Name = kAPIKey, Type = String, Dynamic = False, Default = \"", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kAPISecret, Type = String, Dynamic = False, Default = \"", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kMaxSockets, Type = Double, Dynamic = False, Default = \"5", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kRxEmail, Type = String, Dynamic = False, Default = \"<\?([^@\\s]+@[^@\\s\\.]+\\.[^@\\s>]+)>\?", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kRxEmailName, Type = String, Dynamic = False, Default = \"\\b([^\\\"]*)\\\"\?\\s\\<.*\\>", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kVersion, Type = Double, Dynamic = False, Default = \"1.0", Scope = Protected
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
