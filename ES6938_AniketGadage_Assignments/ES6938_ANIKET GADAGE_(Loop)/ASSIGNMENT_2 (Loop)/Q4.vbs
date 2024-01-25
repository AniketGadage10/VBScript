
Dim vno: vno=12321
     Dim sum,a
     sum=0
       While vno<>0
       a=vno MOD 10
      
       vno=Int(vno/10)
       sum=sum+a
      Wend
      
wscript.echo sum





'Module VBModule
'    Sub Main()
'     Dim vno: vno=12321
'     Dim sum,a
'     sum=0
 '     do WHile vno>0
'       a=vno MOD 10
'          Console.Write(a)
'       vno=Int(vno/10)
'       sum=sum+a
'      Loop
'      
'      Console.Write(" Sum "&sum)
'
 '   End Sub
'End Module