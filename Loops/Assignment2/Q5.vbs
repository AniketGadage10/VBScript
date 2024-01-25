    dim temp,dif,i,pre
   temp=0
   dif=10000
   Dim arr : arr=Array(35,44,99,66,98,76)
     Msgbox (UBOUND(arr))
	Msgbox (LBOUND(arr))

    for i=0 to UBound(arr)
        if(arr(i)>temp) Then
            temp=arr(i)
         end if
    Next
    
       for i=0 to UBound(arr)
           s=temp-arr(i)
        if(s<dif AND arr(i)<>temp) Then
            pre=arr(i)
	    dif=s
         end if
        Next
    Msgbox (" 1st Highest Element In Array Is "&temp)
     Msgbox (" 2nd Highest Element In Array Is "&pre)
    