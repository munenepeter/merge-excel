<?php


namespace MergeExcel;

use MergeExcel\Parser;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Helper\ProgressBar;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;

class ConvertCommand extends Command {
    protected static $defaultName = 'convert';

    protected function configure() {
        $this
            ->setDescription('Convert a workbook to another format.')
            ->addOption('workbook', null, InputOption::VALUE_REQUIRED, 'Path to the workbook')
            ->addOption('from', null, InputOption::VALUE_OPTIONAL, 'Format to convert from')
            ->addOption('to', null, InputOption::VALUE_OPTIONAL, 'Format to convert to');
    }

    protected function execute(InputInterface $input, OutputInterface $output): int {
        $workbook = $input->getOption('workbook');
        if (empty($workbook)) {
            $output->writeln('<error>The --workbook argument is required for the convert operation.</error>');
            return Command::FAILURE;
        }

        $from = $input->getOption('from');
        $to = $input->getOption('to');

        $output->writeln("Converting workbook: $workbook");

        $parser = new Parser();
        $spreadsheet = IOFactory::load($workbook);
        $sheet_count = $spreadsheet->getSheetCount();

        $progressBar = new ProgressBar($output, $sheet_count);
        $progressBar->setFormat('<info>%current%/%max% [%bar%] %percent:3s%%</info>');
        $progressBar->start();

        // Call the convert method with a progress callback
        $parser->convert($workbook, $from, $to, function ($progress) use ($progressBar) {
            $progressBar->setProgress($progress);
        });

        $progressBar->finish();

        $output->writeln("\nConversion completed!");

        return Command::SUCCESS;
    }
}
