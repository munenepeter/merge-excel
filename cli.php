<?php



use MergeExcel\Parser;
use Symfony\Component\Console\Helper\ProgressBar;
use Symfony\Component\Console\Output\ConsoleOutput;

ini_set('memory_limit', '-1'); //# might over use the memory, so remove the limit
date_default_timezone_set('Africa/Nairobi');

require 'vendor/autoload.php';

if ($argc < 2) {
    display_help("Error: No command provided.");
    exit(1);
}

$command_name = $argv[1];
if (!in_array($command_name, ['convert', 'merge'])) {
    display_help("Error: Unknown command '$command_name'.");
    exit(1);
}

// remove the command name from arguments
array_shift($argv);


switch ($argv[0]) {
    case 'merge':
        if (!isset($argv[1]) ||  $argv[1] === '') {
            display_help_for_merge("Error: Missing path to workbooks to merge.");
            exit(1);
        }

        $path = null;

        if (strpos($argv[1], '--path=') !== 0) {
            display_help_for_merge("Error: Invalid argument format. Expected '--path=uploads");
            exit(1);            
        } 

        $path = substr($argv[1], strlen('--path='));

        $parser = new Parser();
        $output = new ConsoleOutput();

        $workbook_count = count(array_diff(scandir($path), ['..', '.']));

        $progressBar = new ProgressBar($output, $workbook_count);
        $progressBar->setFormat('<info>%current%/%max% [%bar%] %percent:3s%%</info>');
        $progressBar->start();

        $parser->merge($path, function ($progress) use ($progressBar) {
            $progressBar->setProgress($progress);
        });

        $progressBar->finish();
        $output->writeln("\nMerge completed!");
        break;

    case 'convert':
        # code...
        break;
    default:
        # code...
        break;
}


// Run the application
// $application->run();