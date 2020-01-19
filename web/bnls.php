<?php
// Dynamically retrieve the list of online BNLS servers from BNETDocs:
$url = 'https://bnetdocs.org/servers.json?type_id=4&status=1';

header('Content-Type: text/plain;charset=utf-8');

function exit_here($message, $code = 500) {
  if (function_exists('http_response_code')) {
    http_response_code($code); // introduced in PHP 5.4
  } else {
    header($_SERVER['SERVER_PROTOCOL'] . ' ' . $code);
  }
  die($message . "\n");
}

if (!($data = file_get_contents($url))) {
  exit_here(sprintf('Failed to download server list from: %s', $url));
}

if (!($json = json_decode($data, true))) {
  exit_here(sprintf('JSON error #%d while decoding list', json_last_error()));
}

echo "\n"; // not sure why this is here - @carlbennett
foreach ($json['servers'] as $server) {
  echo $server['address'] . "\n";
}
