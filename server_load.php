<?php












function get_server_cpu_usage(){
 
	$load = sys_getloadavg();
	return $load[0];
 
}



function get_server_memory_usage(){
 
	$free = shell_exec('free');
	$free = (string)trim($free);
	$free_arr = explode("\n", $free);
	$mem = explode(" ", $free_arr[1]);
	$mem = array_filter($mem);
	$mem = array_merge($mem);
	$memory_usage = $mem[2]/$mem[1]*100;
 
	return $mem[2].' di '.$mem[1].'='.$memory_usage.'%';
}





?>


<p><span class="description">Server Memory Usage:</span> <span class="result"><?= get_server_memory_usage() ?></span></p>
<p><span class="description">Server CPU Usage: </span> <span class="result"><?= get_server_cpu_usage() ?>%</span></p>