<?php

namespace MergeExcel;

use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Helper\ProgressBar;
use Symfony\Component\Console\Attribute\AsCommand;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;

#[AsCommand(name: 'merge', description: 'Merge workbooks')]
class MergeCommand extends Command {


    public function __construct() {
        parent::__construct('merge');
    }
    protected function configure():void {
        $this
            ->setName('merge') 
            ->setDescription('Merge workbooks.')
            ->addOption('path', null, InputOption::VALUE_REQUIRED, 'Path to the workbooks');
    }

    protected function execute(InputInterface $input, OutputInterface $output): int {
        $path = $input->getOption('path');
        if (empty($path)) {
            $output->writeln('<error>The --path argument is required for the merge operation.</error>');
            return Command::FAILURE;
        }

        $output->writeln("Merging workbooks at path: $path");

        $parser = new Parser();

        $workbook_count = count(array_diff(scandir($path), ['..', '.']));

        $progressBar = new ProgressBar($output, $workbook_count);
        $progressBar->setFormat('<info>%current%/%max% [%bar%] %percent:3s%%</info>');
        $progressBar->start();

        $parser->merge($path, function ($progress) use ($progressBar) {
            $progressBar->setProgress($progress);
        });

        $progressBar->finish();
        $output->writeln("\nMerge completed!");

        return Command::SUCCESS;
    }
}
