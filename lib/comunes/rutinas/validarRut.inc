<script Language="VBScript">

function ValidaRut(rut)
if len(rut.value) > 0 then
  dim a,dv,c,s,rt     
  a="NO"                    
  dv=""                     
  c=2                       
  s=0                       
  rt=0                      
  r = replace(rut.value,"-","")
  dvconc = mid(r,len(r),1)
  rutCont = mid(r,1,len(r)-1)

  r = trim(rutCont+"-"+dvconc)              

  if  len(r) > 7 then
      r = ucase(r)
      r = replace(r,".","")
      dv = mid(r,len(r),1)
      r = mid(r,1,len(r)-2)
      for x = 1 to len(r)
          if asc(mid(r,x,1)) > 57 or asc(mid(r,x,1)) < 48 then
            a="SI"
          end if
      next

      If a = "NO" Then
         For x = len(r) to 1 step - 1
             s = s + ( cint( mid(r,x,1) ) * c )
             c = c + 1
             if c > 7 then c = 2
         next  
         rt = ( s mod 11) * 10
         rt = ( rt mod 11 )    
  
         if rt = 10 then rt="K"
         if Trim(cstr(dv)) <> Trim(cstr(rt)) then a="SI"
     End if

     If a = "SI" Then   
      //alert("Rut Mal Ingresado")  
      call mensaje_rut_invalido()
  
     Else  
      rut.value = cstr(r)+"-"+ cstr(rt)
     End if 
  ELSE
    //alert("Rut Mal Ingresado") 
    call mensaje_rut_invalido()

  end if
end if
End Function


</SCRIPT>


<script type="text/javascript">
  function mensaje_rut_invalido()
  {
    $('#span_RUT_USUARIO').text("RUT invalido")
    $('#RUT_USUARIO').css('border-color','#FE2E2E') 
  }
</script>