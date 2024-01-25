
------------------------------------------------------------------------------------------

'                             Online VB Compiler.
'                 Code, Compile, Run and Debug VB program online.
' Write your code in this editor and press "Run" button to execute it.

Q3

Module VBModule
    Sub Main()
     Dim vno: vno=5
      Dim i,a,x,m,s,z
        x=1
      for i=1 to vno 
        s=vno-i
        
         for a=1 to s
             CONSOLE.Write(" ")
         Next
      
 
        for m=1 to i
             CONSOLE.Write(m)
        Next 
        
         m=m-2
         
        for z=1 to i-1
            CONSOLE.Write(m)
            m=m-1
        Next 
     
        CONSOLE.Writeline()
        x=x+2
        
    Next
      
    End Sub
End Module
------------------------------------------------------


