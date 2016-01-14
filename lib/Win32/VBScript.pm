package Win32::VBScript;

use strict;
use warnings;

use Carp;
use Digest::SHA qw(sha1_hex);
use File::Slurp;

require Exporter;
our @ISA         = qw(Exporter);
our %EXPORT_TAGS = ('all' => [qw(
    compile_vbs compile_js cscript wscript
)]);
our @EXPORT      = qw();
our @EXPORT_OK   = ( @{ $EXPORT_TAGS{'all'} } );

my $VBRepo = $ENV{'TEMP'}.'\\Repo01';

sub new {
    my $pkg = shift;

    my ($lang, $code) = @_;

    unless ($lang eq 'vbs' or $lang eq 'js') {
        croak "E010: Invalid language ('$lang'), expected ('vbs' or 'js')";
    }

    unless (-d $VBRepo) {
        mkdir $VBRepo or croak "E020: Can't mkdir '$VBRepo' because $!";
    }

    my $clear = '';

    for (@$code) {
        $clear .= $_."\n";
    }

    my $sha1 = sha1_hex($clear);

    my $fname = $lang.'-'.$sha1.'.'.$lang;
    my $full = $VBRepo.'\\'.$fname;

    unless (-f $full) {
        write_file($full, $clear);
    }

    bless {
      'fn'   => $fname,
      'lang' => $lang,
    }, $pkg;
}

sub compile_vbs {
    my ($code) = @_;
    Win32::VBScript->new('vbs', $code);
}

sub compile_js {
    my ($code) = @_;
    Win32::VBScript->new('js', $code);
}

sub _run {
    my $self = shift;
    my ($scr) = @_;

    unless ($scr eq 'cscript' or $scr eq 'wscript') {
        croak "E030: Invalid script ('$scr'), expected ('cscript' or 'wscript')";
    }

    my $fname = $self->{'fn'};
    my $full  = $VBRepo.'\\'.$fname;
    my $lang  = $self->{'lang'};

    unless (-f $full) {
        croak "E040: Panic -- can't find executable '$full'";
    }

    my $engine =
      $lang eq 'vbs' ? 'VBS'     :
      $lang eq 'js'  ? 'JScript' :
      croak "E050: Panic -- invalid language ('$lang'), expected ('vbs' or 'js')";

    system qq{$scr //Nologo //E:$engine "$full"};
}

sub cscript {
    my $self = shift;
    $self->_run('cscript');
}

sub wscript {
    my $self = shift;
    $self->_run('wscript');
}

1;

__END__

=head1 NAME

Win32::VBScript - Run Visual Basic programs

=head1 SYNOPSIS

    use strict;
    use warnings;

    use Win32::VBScript qw(:all);

    # This is the procedural interface:
    # *********************************

    my $p1 = compile_js ([ qq{WScript.StdOut.WriteLine("Bonjour");} ]); cscript($p1);
    my $p2 = compile_vbs([ qq{WScript.StdOut.WriteLine "Hello"}     ]); cscript($p2);

    # This is the OO interface:
    # *************************

    compile_js ([ qq{WScript.StdOut.WriteLine("Test1");} ])->cscript;
    compile_vbs([ qq{WScript.StdOut.WriteLine "Test2"}   ])->cscript;

    # And with vbs, of course, you can use MsgBox:
    # ********************************************

    compile_vbs([ qq{MsgBox "Greetings Earthlings..."} ])->wscript;

=head1 DESCRIPTION

This module allows you to invoke code fragments written in Visual
Basic (or even JavaScript) from within a perl program.

=head1 AUTHOR

Klaus Eichner <klaus03@gmail.com>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2016 by Klaus Eichner

All rights reserved. This program is free software; you can redistribute
it and/or modify it under the terms of the artistic license 2.0,
see http://www.opensource.org/licenses/artistic-license-2.0.php

=cut
