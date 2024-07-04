<?php



use MergeExcel\ConvertCommand;
use Symfony\Component\Console\Application;

ini_set('memory_limit', '-1'); //# might over use the memory, so remove the limit
date_default_timezone_set('Africa/Nairobi');

require 'vendor/autoload.php';

$application = new Application();


$application->add(new \MergeExcel\MergeCommand());
$application->add(new ConvertCommand());

if ($argc < 2) {
    echo "Error: No command provided.\n";
    echo "Usage:\n";
    echo "  php index.php convert --workbook=WORKBOOK [--from=FORMAT] [--to=FORMAT]\n";
    echo "  php index.php merge --path=PATH\n";
    exit(1);
}

$commandName = $argv[1];
if (!in_array($commandName, ['convert', 'merge'])) {
    echo "Error: Unknown command '$commandName'.\n";
    echo "Usage:\n";
    echo "  php index.php convert --workbook=WORKBOOK [--from=FORMAT] [--to=FORMAT]\n";
    echo "  php index.php merge --path=PATH\n";
    exit(1);
}

// Remove the command name from arguments
array_shift($argv);

// Run the application
$application->run();