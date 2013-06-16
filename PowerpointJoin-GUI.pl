#!/usr/bin/env perl

use strict;
use Carp;
use Wx;
use Win32::PowerpointJoin;

my $ROW = 30;

package MyFrame;
use Wx::Event qw(EVT_BUTTON);
use base qw(Wx::Frame);
sub new {
    my $class = shift;
    my $self = $class->SUPER::new(@_);
    my $panel = Wx::Panel->new($self, -1);
    $self->{panel} = $panel;

    $self->{txt_select} = Wx::StaticText->new(  $panel,
                                                123,
                                                "Select a configuration file",
                                                [10, $ROW]
                                             );

    $self->SetConfig("(no configuration file selected)");

    my $BTNID_BROWSE = 200;
    $self->{btn_browse} = Wx::Button->new(  $panel,
                                            $BTNID_BROWSE,
                                            "...",
                                            [150,$ROW-2],
                                            [30, 20],
                                        );
    EVT_BUTTON( $self, $BTNID_BROWSE, \&Browse);

    $self->{filepicker} = Wx::FileDialog->new(  $panel,
                                                "Browse",
                                                $ENV{USERPROFILE} . "\\Documents",
                                                '',
                                                '*.*',
                                             );

    $self->{txt_output} = Wx::StaticText->new(  $panel,
                                                123,
                                                "Type the name of the merged charts",
                                                [10, $ROW*3]
                                             );
    my $TEXTID_OUTPUT = 300;
    $self->{text_output} = Wx::TextCtrl->new( $panel,
                                                $TEXTID_OUTPUT,
                                                'merged.pptx',
                                                [198, ($ROW*3)-2],
                                                [150, 20],
                                                0
                                            );
    my $BTNID_MERGE = 100;
    $self->{btn_merge} = Wx::Button->new(   $panel,
                                            $BTNID_MERGE,
                                            "Merge",
                                            [10, $ROW*4],
                                        );
    EVT_BUTTON ( $self, $BTNID_MERGE, \&Merge);
    return $self;
}

sub SetConfig {
    my ($self, $filename) = @_;
    $self->{txt_selected_config} = Wx::StaticText->new(  $self->{panel},
                                                124543,
                                                $filename,
                                                [10, $ROW*2]
                                             );

}

sub Browse {
    my ($self, $evt) = @_;
    my $picker = $self->{filepicker};
    $picker->ShowModal();
    $self->{config_dir}  = $picker->GetDirectory();
    $self->{config_file} = $picker->GetFilename();
    $self->{config_filename} = $self->{config_dir} . '\\' . $self->{config_file};
    print "Config file: " . $self->{config_filename} . "\n";
    $self->SetConfig($self->{config_filename});
}

sub Merge {
    my ($self, $evt) = @_;
    my $config_file = $self->{config_filename};
    my $output      = $self->{config_dir} . "\\" . "output.pptx";
    Win32::PowerpointJoin::merge($config_file, $output);
}

package Gui;
use base qw(Wx::App);
sub OnInit {
    my ($self) = @_;
    print "creating the frame\n";
    my $frame = MyFrame->new( undef,
                                -1,
                                'Powerpoint Join',
                                [1,1],
                                [400, 250],
                               );
    print "setTopWindow\n";
    $self->SetTopWindow($frame);
    print "show\n";
    $frame->Show(1);
    return 1;
}

package main;
my $wx = Gui->new();
$wx->MainLoop;
