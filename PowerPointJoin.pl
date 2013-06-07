#!/usr/bin/env perl

use strict;
use Carp;
use Data::Dumper;
use Getopt::Long;
use autodie qw(open close);
use FindBin qw($Bin);
use Win32::PowerPoint;
use Win32::OLE::Const 'Microsoft PowerPoint';

our $AUTOLOAD;

my ($HELP, $CONFIG, $OUTPUT);
my $opts = GetOptions(
    'help=s'   => \$HELP,
    'config=s' => \$CONFIG,
    'output=s' => \$OUTPUT,
);

if ($HELP || !($CONFIG && $OUTPUT) ) {
    my $msg =<<EOF;
Usage: $0 --config <config_file> --output <output.ppt>


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
 
my %config;
open my $IN, '<', $CONFIG;
my @lines = <$IN>;
close $IN;
my $current_file;
for(my $i=0; $i<@lines; $i++) {
    my $line = $lines[$i];
    chomp $line;
    next if $line =~ /^$/ || $line =~ /^[#;]/;
    if ($line =~ /^file=(.+)$/) {
        $current_file = $1;
        $config{$current_file}{start} = 1;
        $config{$current_file}{end} = 9999;
    } elsif ( $line =~ /^slides=(.+)$/) {
        $config{$current_file}{slides} = $1;
    } else {
        print "Unrecognized line in $CONFIG: $line\n";
        exit;
    }
}


sub dump_ole {
    my ($name, $obj) = @_;
    print "Properties: $name\n--------------------\n";
    my @k = sort keys %{$obj};
    foreach my $key (sort keys %{$obj}) {
        my $value;
        eval { $value = $obj->{$key} };
        $value = "***Exception: $@";
        $value = "<undef>" unless defined $value;
        $value = '[' . Win32::OLE->QueryObjectType($value) . ']'
            if UNIVERSAL::isa($value, 'Win32::OLE');

        $value = '(' . join(',', @$value) . ')' if ref($value) eq 'ARRAY';
        printf "%s %s %s\n", $key, '.' x (40-length($key)), $value;
    }
    print "\nMethods: $name\n----------------------\n";

    my $typeinfo = $obj->GetTypeInfo();
    my $attr = $typeinfo->_GetTypeAttr();
    my @functions;
    for (my $i = 0; $i< $attr->{cFuncs}; $i++) {
        my $desc = $typeinfo->_GetFuncDesc($i);
        # the call conversion of method was detailed in %$desc
        my $funcname = @{$typeinfo->_GetNames($desc->{memid}, 1)}[0];
        push(@functions, $funcname);
    }
    print join("\n  ->", sort @functions);
    print "\n\n";
}

sub AUTOLOAD {
    my $obj = shift;
    $AUTOLOAD =~ s/^.*:://;
    my $meth = $AUTOLOAD;
    $AUTOLOAD = "SUPER::" . $AUTOLOAD;
    my $retval = $obj->AUTOLOAD(@_);
    unless (defined($retval) || $AUTOLOAD eq 'DESTROY') {
        my $err = Win32::OLE::LastError();
        croak(sprintf("$meth returnled OLE error 0x%08x", $err))
            if ($err);
        return $retval;
    }
}

# $ole
# $ole->Presentations->Add
# $ppt = $ole->Presentations->Open(filename)
# check file exists
# $ppt->Slides

my $PPT_STRING = 'Powerpoint.Application';
my $pp = Win32::PowerPoint->new;

# Get a count of the slides
foreach my $key ( sort keys %config) {
    my $file = "$Bin\\$key";
    $file =~ s{\\}{\\\\}g;
    $file =~ s{\/}{\\\\}g;
    my $ppt_insert = $pp->application->Presentations->Open($file);
    my $n_slides = $ppt_insert->Slides->Count;
    $ppt_insert->Close;
    $config{$key}{count} = $n_slides;
    print "$file ($n_slides slides)\n";
}

$pp->new_presentation;
my $ppt = $pp->presentation;

foreach my $key ( sort keys %config ) {
    my $file = "$Bin\\$key";
    $file =~ s{\\}{\\\\}g;
    $file =~ s{\/}{\\\\}g;

    my $count = $config{$key}{count}; 

    print "File=$file ($count slides)\n";
    # app.Slides.InsertFromFile("filename", "index", "slideStart", "slideend")
    
    my $last = $ppt->Slides->Count;
    if (! exists $config{$key}{slides} ) {
        print "   > inserting all (1-$count)\n";
        my $n_inserted = $ppt->Slides->InsertFromFile($file, $last, 1, $count);
        if (! $n_inserted) {
        }
        if ( ! $n_inserted ) {
            print "Error: No slides were inserted: " . Win32::OLE->LastError() . "\n";
            print "File:  $file\n";
            print "Count: $count\n";
            print "Range: all\n";
            exit;
        }
    } else {
        my $slides = $config{$key}{slides};
        $slides =~ s/\s+//g;
        my @ranges = split(/[;,]/, $slides);
        foreach my $range (@ranges) {

            my ($start, $end); 
            if ($range =~ /^(\d+)-(\d+)$/) {
                ($start, $end) = ($1, $2);
            } elsif ($range =~ /^\d+$/) {
                ($start, $end) = ($range, $range);
            } else {
                print "Error: Invalid range\n";
                print "File:  $file\n";
                print "Count: $count\n";
                print "Range: $range\n";
                exit;
            }
            if ($start > $end || $start > $count || $end > $count) {
                print "Error: Invalid range\n";
                print "File:  $file\n";
                print "Count: $count\n";
                print "Start: $start\n";
                print "End:   $end\n";
                exit;
            }
            my $n_inserted = $ppt->Slides->InsertFromFile($file, $last, $start, $end);
            print "   > inserted $n_inserted slides (range $range)\n";
            if ( ! $n_inserted ) {
                print "Error: No slides were inserted: " . Win32::OLE->LastError() . "\n";
                print "File:  $file\n";
                print "Count: $count\n";
                print "Range: $range\n";
                print "Start: $start\n";
                print "End:   $end\n";
                exit;
            }
        } 
    }

        
}

$pp->save_presentation($OUTPUT);
$pp->close_presentation;
undef $pp;

if (! -f $OUTPUT) {
    print "$OUTPUT did not get saved\n";
}

