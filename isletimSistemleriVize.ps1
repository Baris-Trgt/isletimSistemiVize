function Cikti{

$servis = @{}
$port = @{}

$srvs =  Get-WmiObject win32_service | Select-Object  Name , ProcessId 
  
$prt =  netstat -a -n -o | ConvertFrom-String | Select-Object p2,p3,p4,p5,p6


 foreach($i in $srvs)
 {
    foreach($j in $prt)
      { 
            
            if($i.ProcessId -eq $j.p5)
                {
                    
                   $i.Name | %{
                   
                        if(!$servis.Contains($_))
                            {
                            
                              $servis.Add($_,$null)
                              'Servis = ' +   $_   
                            }
                    }


                   $j.p3.Split(".,[]:")[-1] | %{
                   
                        if(!$port.Contains($_))
                            {
                            
                              $port.Add($_,$null)
                              'Port = ' + $_ 
                            }
                    }

                }
     }

     
}

}

Cikti > C:\Windows\Temp\Odev.txt