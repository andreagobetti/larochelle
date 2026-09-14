
<?php
echo(CP1251toUTF8("città"));


function CP1251toUTF8($string){ 
  $out = ''; 
  for ($i = 0; $i<strlen($string); ++$i){ 
    $ch = ord($string{$i}); 
    if ($ch < 0x80) $out .= chr($ch); 
    else 
      if ($ch >= 0xC0) 
        if ($ch < 0xF0) 
             $out .= "\xD0".chr(0x90 + $ch - 0xC0); // &#1040;-&#1071;, &#1072;-&#1087; (A-YA, a-p) 
        else $out .= "\xD1".chr(0x80 + $ch - 0xF0); // &#1088;-&#1103; (r-ya) 
      else 
        switch($ch){ 
          case 0xA8: $out .= "\xD0\x81"; break; // YO 
          case 0xB8: $out .= "\xD1\x91"; break; // yo 
          // ukrainian 
          case 0xA1: $out .= "\xD0\x8E"; break; // &#1038; (U) 
          case 0xA2: $out .= "\xD1\x9E"; break; // &#1118; (u) 
          case 0xAA: $out .= "\xD0\x84"; break; // &#1028; (e) 
          case 0xAF: $out .= "\xD0\x87"; break; // &#1031; (I..) 
          case 0xB2: $out .= "\xD0\x86"; break; // I (I) 
          case 0xB3: $out .= "\xD1\x96"; break; // i (i) 
          case 0xBA: $out .= "\xD1\x94"; break; // &#1108; (e) 
          case 0xBF: $out .= "\xD1\x97"; break; // &#1111; (i..) 
          // chuvashian 
          case 0x8C: $out .= "\xD3\x90"; break; // &#1232; (A) 
          case 0x8D: $out .= "\xD3\x96"; break; // &#1238; (E) 
          case 0x8E: $out .= "\xD2\xAA"; break; // &#1194; (SCH) 
          case 0x8F: $out .= "\xD3\xB2"; break; // &#1266; (U) 
          case 0x9C: $out .= "\xD3\x91"; break; // &#1233; (a) 
          case 0x9D: $out .= "\xD3\x97"; break; // &#1239; (e) 
          case 0x9E: $out .= "\xD2\xAB"; break; // &#1195; (sch) 
          case 0x9F: $out .= "\xD3\xB3"; break; // &#1267; (u) 
        } 
  } 
  return $out; 
} 
?>
		
