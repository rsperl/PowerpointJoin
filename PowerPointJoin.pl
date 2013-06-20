#!/usr/bin/env perl

use strict;
use Carp;
use Getopt::Long;
use FindBin qw($Bin);
use Win32::PowerpointJoin;

my ($HELP, $CONFIG, $OUTPUT, $START_WITH);
my $opts = GetOptions(
    'help=s'       => \$HELP,
    'config=s'     => \$CONFIG,
    'start-with=s' => \$START_WITH,
    'output=s'     => \$OUTPUT,
);

if ($HELP || !$CONFIG && !($OUTPUT || $START_WITH) ) {
    my $msg =<<EOF;
Usage: $0 --config <config_file> [--start-with <charts.pptx> | --output <output.ppt>]


The config file should have the following format:

file=filename1.ppt
slides=1-4,6-8

file=filename2.ppt

file=filename3.ppt
slides=3

The start and end lines are optional. start defaults to 1 and end
defaults to the last slide in the file. The blank line betwen
sections is option, but makes it readable if you have a lot of
start/end lines.
EOF
    print $msg;
    exit;
}

if($START_WITH) {
    $START_WITH = "$Bin\\$START_WITH";
    die "Existing charts $START_WITH does not exist\n" unless -r $START_WITH;
}

&Win32::PowerpointJoin::merge($CONFIG, { start_with => $START_WITH, output => $OUTPUT});
