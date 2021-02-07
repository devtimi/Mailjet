#tag Class
Protected Class MailJet
	#tag Method, Flags = &h21
		Private Function ConvertEmailToMailJet(_oMail as EmailMessage) As Dictionary
		  // Validate From
		  if _oMail.FromAddress.Trim = "" then
		    var _ex as new MailJetException
		    _ex.Message = "EmailMessage has no FromAddress"
		    RaiseEvent Error(_ex)
		    return nil
		    
		  end
		  
		  // Validate To
		  if _oMail.ToAddress.Trim = "" then
		    var _ex as new MailJetException
		    _ex.Message = "EmailMessage has no Recepients"
		    RaiseEvent Error(_ex)
		    return nil
		    
		  end
		  
		  var _ardictFrom() as Dictionary = GetAddressArray(_oMail.FromAddress)
		  
		  var _ardictTo() as Dictionary = GetAddressArray(_oMail.ToAddress)
		  var _ardictCC() as Dictionary = GetAddressArray(_oMail.CCAddress)
		  var _ardictBCC() as Dictionary = GetAddressArray(_oMail.BCCAddress)
		  
		  // Start building JSON item
		  var _dictBody as new Dictionary
		  _dictBody.Value("From") = _ardictFrom(0)
		  _dictBody.Value("To") = _ardictTo
		  
		  if _ardictCC.LastIndex > -1 then
		    _dictBody.Value("CC") = _ardictCC
		    
		  end
		  
		  if _ardictBCC.LastIndex > -1 then
		    _dictBody.Value("BCC") = _ardictBCC
		    
		  end
		  
		  _dictBody.Value("Subject") = _oMail.Subject
		  
		  if _oMail.BodyPlainText <> "" then
		    _dictBody.Value("TextPart") = _oMail.BodyPlainText
		    
		  end
		  
		  if _oMail.BodyHTML <> "" then
		    _dictBody.Value("HTMLPart") = _oMail.BodyHTML
		    
		  end
		  
		  // Attachments are not handled at this time
		  if _oMail.Attachments.LastIndex > -1 then
		    break
		    
		  end
		  
		  return _dictBody
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetAddressArray(_sRecipientsCSV as String) As Dictionary()
		  var _ardictTo() as Dictionary
		  
		  var _arsAddressed() as String = _sRecipientsCSV.Split(",")
		  
		  for each _sAddress as String in _arsAddressed
		    var _dictTo as new Dictionary
		    
		    var _rx as new RegEx
		    _rx.SearchPattern = kRxEmail
		    
		    var _rxm as RegExMatch = _rx.Search(_sAddress)
		    if _rxm <> nil then
		      _dictTo.Value("Email") = _rxm.SubExpressionString(1)
		      
		    end
		    
		    // Now check for name
		    _rx = new RegEx
		    _rx.SearchPattern = kRxEmailName
		    _rxm = _rx.Search(_sAddress)
		    
		    if _rxm <> nil then
		      _dictTo.Value("Name") = _rxm.SubExpressionString(1)
		      
		    end
		    
		    // Add it to the array
		    _ardictTo.Add(_dictTo)
		    
		  next _sAddress
		  
		  return _ardictTo
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetFromObject(_oMail as EmailMessage) As Dictionary
		  // Set up From object
		  var _dictFrom as new Dictionary
		  
		  var _rx as new RegEx
		  _rx.SearchPattern = kRxEmail
		  
		  var _rxm as RegExMatch = _rx.Search(_oMail.FromAddress)
		  if _rxm <> nil then
		    _dictFrom.Value("Email") = _rxm.SubExpressionString(1)
		    
		  end
		  
		  // Now check for name
		  _rx = new RegEx
		  _rx.SearchPattern = kRxEmailName
		  _rxm = _rx.Search(_oMail.FromAddress)
		  
		  if _rxm <> nil then
		    _dictFrom.Value("Name") = _rxm.SubExpressionString(1)
		    
		  end
		  
		  return _dictFrom
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleError(_oSender as URLConnection, ex as RuntimeException)
		  #pragma unused _oSender
		  
		  // Raise Error event.
		  Error(ex)
		  
		  mbBusy = false
		  moSock = nil
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleResponse(_oSender As URLConnection, URL As String, HTTPStatus As Integer, content As String)
		  #pragma unused _oSender
		  
		  if HTTPStatus <> 200 then
		    var _ex as new MailJetException
		    _ex.Message = "HTTP response code was not okay: " + HTTPStatus.ToString
		    
		    try
		      // Attempt to parse out the error messager
		      var _vResponse as Variant = ParseJSON(content.DefineEncoding(Encodings.UTF8))
		      var _dictResponse as Dictionary = Dictionary(_vResponse)
		      
		      if _dictResponse.HasKey("ErrorMessage") then
		        _ex.Message = _dictResponse.Value("ErrorMessage")
		        
		      end
		      
		    catch ex as IllegalCastException
		      // Trying to turn the variant into a dictionary failed
		      // This can happen if it's not the single json object we're expecting
		      
		    catch ex as InvalidJSONException
		      // Response wasn't json
		      // Not much we can parse here
		      
		    end try
		    
		    RaiseEvent Error(_ex)
		    
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
		Private Sub HandleResponse200(_sBody as String)
		  var _vBody as Variant
		  
		  try
		    _vBody = ParseJSON(_sBody)
		    
		  catch ex as InvalidJSONException
		    // Response wasn't json
		    // Not much we can parse here
		    RaiseEvent Error(ex)
		    return
		    
		  end try
		  
		  break
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function NewSocket() As URLConnection
		  // Create URLConnection and add the authentication header
		  var _sAuth as String = EncodeBase64(kAPIKey + ":" + kAPISecret, 0)
		  
		  var _oSock as new URLConnection
		  _oSock.RequestHeader("Authorization") = "Basic " + _sAuth
		  
		  // Handle server responses that aren't 200
		  AddHandler _oSock.Error, WeakAddressOf HandleError
		  AddHandler _oSock.ContentReceived, WeakAddressOf HandleResponse
		  
		  return _oSock
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
		    var _ex as new MailJetException
		    _ex.Message = "This MailJet socket is already in use, please wait for the MailSent event."
		    RaiseEvent Error(_ex)
		    return
		    
		  end
		  
		  var _dictRequest as new Dictionary
		  
		  #if DebugBuild then
		    // This SandboxMode flag prevents emails from actually sending
		    // Take this out if you're testing actual delivery
		    _dictRequest.Value("SandboxMode") = true
		    
		  #endif
		  
		  var _aroMessages() as Dictionary
		  
		  // Put all messages into request
		  for each _oEmail as EmailMessage in Messages
		    var _dictMail as Dictionary = ConvertEmailToMailJet(_oEmail)
		    _aroMessages.Add(_dictMail)
		    
		  next _oEmail
		  
		  _dictRequest.Value("Messages") = _aroMessages()
		  
		  var _sBody as String = GenerateJSON(_dictRequest)
		  
		  mbBusy = true
		  moSock = NewSocket
		  
		  
		  moSock.SetRequestContent(_sBody, "application/json")
		  
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

	#tag Constant, Name = kRxEmail, Type = String, Dynamic = False, Default = \"<\?([^@\\s]+@[^@\\s\\.]+\\.[^@\\.\\s>]+)>\?", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kRxEmailName, Type = String, Dynamic = False, Default = \"(.*)\\s\\<.*\\>", Scope = Private
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
