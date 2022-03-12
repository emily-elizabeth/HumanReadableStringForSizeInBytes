#tag Module
Protected Module HumanReadableStringForSizeInBytes
	#tag Method, Flags = &h0
		Function HumanReadableStringForSizeInBytes(byteCount As UInt64, formatterStyle As ByteCountFormatterStyle = ByteCountFormatterStyle.Decimal) As String
		  DIM returnValue As String
		  
		  DIM kilobyte As UInt64 = if(formatterStyle = ByteCountFormatterStyle.Decimal, 1000, 1024)
		  
		  if (byteCount = 1) then
		    returnValue = Replace(oneByte, "%number%", "1")
		  elseif (byteCount < kilobyte) then
		    returnValue = Replace(numberBytes, "%number%", Str(byteCount))
		  else
		    DIM exp As Integer = (Log(byteCount) / Log(kilobyte))
		    DIM sizeLabel As String = Mid("KMGTPE", exp, 1) + if(formatterStyle = ByteCountFormatterStyle.Decimal, "B", "iB")
		    returnValue = Format((byteCount / (kilobyte ^ exp)), "-#.0") + " " + sizeLabel
		  end if
		  
		  Return returnValue
		End Function
	#tag EndMethod


	#tag Constant, Name = numberBytes, Type = String, Dynamic = False, Default = \"", Scope = Private
	#tag EndConstant

	#tag Constant, Name = oneByte, Type = String, Dynamic = False, Default = \"", Scope = Private
	#tag EndConstant


	#tag Enum, Name = ByteCountFormatterStyle, Flags = &h0
		Decimal = 2
		Binary = 3
	#tag EndEnum


End Module
#tag EndModule
