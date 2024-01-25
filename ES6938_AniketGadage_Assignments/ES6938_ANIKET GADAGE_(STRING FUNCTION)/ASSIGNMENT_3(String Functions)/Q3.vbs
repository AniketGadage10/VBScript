str1="Expleo"

str2="India"

temp=str1

Wscript.echo " Before Swap : String1 = "&str1&" String2 = "&str2

str1=Replace(str1,str1,str2)

str2=Replace(str2,str2,temp)

Wscript.echo " After Swap : String1 = "&str1&" String2 = "&str2