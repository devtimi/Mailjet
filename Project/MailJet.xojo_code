#tag Class
Protected Class MailJet
	#tag Method, Flags = &h0
		Sub Constructor()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ConvertEmailToMailJet(_oMail as EmailMessage) As Dictionary
		  var _dictBody as new Dictionary
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function NewSocket() As URLConnection
		  // Create URLConnection and add the authentication header
		  var _sAuth as String = EncodeBase64(kAPIKey + ":" + kAPISecret, 0)
		  
		  var _oSock as new URLConnection
		  _oSock.RequestHeader("Authorization") = "Basic " + _sAuth
		  
		  // Retain the request
		  // maroRequests.Add(toSock)
		  
		  // Handle server responses that aren't 200
		  // AddHandler toSock.Error, WeakAddressOf HandleError
		  // AddHandler toSock.ServerResponse, WeakAddressOf HandleServerResponse
		  
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
		  var _dictRequest as new Dictionary
		  
		  #if DebugBuild then
		    _dictRequest.Value("SandboxMode") = true
		    
		  #endif
		  
		  var _aroMessages() as Dictionary
		  
		  // Put all messages into request
		  for each _oEmail as EmailMessage in Messages
		    var _dictMail as Dictionary = ConvertEmailToMailJet(_oEmail)
		    _aroMessages.Add(_dictMail)
		    
		  next _oEmail
		  
		  _dictRequest.Value("Messages") = _aroMessages()
		End Sub
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event MailSent()
	#tag EndHook


	#tag Property, Flags = &h0
		Messages() As EmailMessage
	#tag EndProperty


	#tag Constant, Name = kAPIKey, Type = String, Dynamic = False, Default = \"", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kAPISecret, Type = String, Dynamic = False, Default = \"", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kMaxSockets, Type = Double, Dynamic = False, Default = \"5", Scope = Private
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
