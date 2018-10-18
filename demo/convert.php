<?php
require '../vendor/autoload.php';

use Office2PDF\Generator;

$src = 'source_docs';
$files = scandir($src);
$files = array_filter($files, function($file) {
    return !in_array($file, ['.', '..']);
});
array_walk($files, function(&$file, $key) use($src) {
    $file = $src . '/' . $file;
});

$converter = new Generator($files);
echo $converter->convert('./output') . " files converted.\n";
