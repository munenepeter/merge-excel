<?php

use Symfony\Component\Console\Formatter\OutputFormatterStyle;
use Symfony\Component\Console\Output\ConsoleOutput;

function display_help(string $help = ''): void {

    $output = new ConsoleOutput();


    $styleTitle = new OutputFormatterStyle('white', 'blue', ['bold']);
    if ($help !== '' && str_contains($help, "Error")) {
        $styleTitle = new OutputFormatterStyle('white', 'red', ['bold']);
    }
    $styleHeader = new OutputFormatterStyle('bright-cyan', null, ['bold']);
    $styleOption = new OutputFormatterStyle('yellow');
    $styleDescription = new OutputFormatterStyle('white');


    $output->getFormatter()->setStyle('title', $styleTitle);
    $output->getFormatter()->setStyle('header', $styleHeader);
    $output->getFormatter()->setStyle('option', $styleOption);
    $output->getFormatter()->setStyle('description', $styleDescription);

    $output->writeln("");
    ($help === '') ? $output->writeln("<title>Help</title>") : $output->writeln("<title>$help</title>");
    $output->writeln("<header>Usage:</header>");
    $output->writeln("  <option>php index.php [command] [options]</option>");
    $output->writeln("");
    $output->writeln("<header>Commands:</header>");
    $output->writeln("  <option>merge</option>           <description>Merge workbooks</description>");
    $output->writeln("  <option>convert</option>         <description>Convert a workbook to another format</description>");
    $output->writeln("");
    $output->writeln("<header>Examples:</header>");
    $output->writeln("  <option>php index.php merge --path=/path/to/workbooks</option>");
    $output->writeln("  <option>php index.php convert --workbook=/path/to/workbook --from=xls --to=xlsx</option>");
}


function display_help_for_merge(string $help): void {
    $output = new ConsoleOutput();


    $styleTitle = new OutputFormatterStyle('white', 'blue', ['bold']);

    if (str_contains($help, "Error")) {
        $styleTitle = new OutputFormatterStyle('white', 'red', ['bold']);
    }

    $styleHeader = new OutputFormatterStyle('bright-cyan', null, ['bold']);
    $styleOption = new OutputFormatterStyle('yellow');
    $styleDescription = new OutputFormatterStyle('white');


    $output->getFormatter()->setStyle('title', $styleTitle);
    $output->getFormatter()->setStyle('header', $styleHeader);
    $output->getFormatter()->setStyle('option', $styleOption);
    $output->getFormatter()->setStyle('description', $styleDescription);

    $output->writeln("");
    $output->writeln("<title>$help</title>");
    $output->writeln("<header>Examples:</header>");
    $output->writeln("  <option>php index.php merge --path=/path/to/workbooks</option>");
}
