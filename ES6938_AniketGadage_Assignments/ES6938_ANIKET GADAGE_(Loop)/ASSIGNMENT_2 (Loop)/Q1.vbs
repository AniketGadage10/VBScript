'                             Online VB Compiler.
'                 Code, Compile, Run and Debug VB program online.
' Write your code in this editor and press "Run" button to execute it.

Q1

Module VBModule
    Sub Main()
     Dim vno: vno=6
      Dim i,a,x
        
      for i=1 to vno 
         for a=1 to i
             CONSOLE.Write(a)
         Next
          CONSOLE.Writeline()
     Next
     vno=vno-1
    for i=1 to vno 
          for x = 1 to vno
             CONSOLE.Write(x)
         Next
          CONSOLE.Writeline()
          vno=vno-1
     Next
      
    End Sub
End Module


----------------------------------------------------------------------------------------------
