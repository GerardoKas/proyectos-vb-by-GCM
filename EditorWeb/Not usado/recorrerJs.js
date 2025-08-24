window.onerror='return false'
obj=prompt("Objeto?","document.")
if (obj=="") 
{
alert("Terminando");
}
tobj=eval(obj)
document.open()
document.write ("Objeto: " + obj + "<br>" + "Objeto? : " + tobj + "<br><hr>")

for (i in tobj){

 document.write (i + "=" )   
 document.write (tobj[i] + "<br>\n")
}
alert("Fin")


