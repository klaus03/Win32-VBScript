package Win32::VBScript;

use strict;
use warnings;

use Carp;
use Digest::SHA qw(sha1_hex);
use File::Slurp;
use Win32::OLE;

require Exporter;
our @ISA         = qw(Exporter);
our %EXPORT_TAGS = ('all' => [qw(
    compile_vbs compile_js cscript wscript func
)]);
our @EXPORT      = qw();
our @EXPORT_OK   = ( @{ $EXPORT_TAGS{'all'} } );

my $VBRepo = $ENV{'TEMP'}.'\\Repo01';

sub new {
    my $pkg = shift;

    my ($type, $lang, $code) = @_;

    unless ($type eq 'prog' or $type eq 'func') {
        croak "E010: Invalid type ('$type'), expected ('prog' or 'func')";
    }

    unless (-d $VBRepo) {
        mkdir $VBRepo or croak "E020: Can't mkdir '$VBRepo' because $!";
    }

    my $dat_engine;
    my $dat_comment;

    if ($lang eq 'vbs') {
        $dat_engine  = 'VBScript';
        $dat_comment = q{'};
    }
    elsif ($lang eq 'js') {
        $dat_engine  = 'JScript';
        $dat_comment = q{//};
    }
    else {
        croak "E030: Invalid language ('$lang'), expected ('vbs' or 'js')";
    }

    my $dat_text  = ''; for (@$code) { $dat_text .= $_."\n"; }
    my $dat_sha1  = sha1_hex($dat_text);
    my $dat_class = "InlineWin32COM.WSC\_$dat_sha1.wsc";

    my %dat_func;

    for (split m{\n}xms, $dat_text) {
        if (m{\A \s* (?: function | sub) \s+ (\w+) (?: \z | \W)}xmsi) {
            $dat_func{$1} = undef;
        }
    }

    my $file_content;

    if ($type eq 'prog') {
        $file_content = $dat_comment.' -- '.$dat_engine.qq{\n\n}.$dat_text;
    }
    elsif ($type eq 'func') {
        $file_content =
          qq{<?xml version="1.0"?>\n}.
          qq{<component>\n}.
          qq{  <registration }.
          qq{    description="Inline::WSC Class" }.
          qq{    progid="$dat_class" }.
          qq{    version="1.0">\n}.
          qq{  </registration>\n}.
          qq{  <public>\n}.
          join('', map { qq{    <method name="$_" />\n} } sort { lc($a) cmp lc($b) } keys %dat_func).
          qq{  </public>\n}.
          qq{  <implements type="ASP" id="ASP" />\n}.
          qq{  <script language="$dat_engine">\n}.
          qq{    <![CDATA[\n$dat_text\n]]>\n}.
          qq{  </script>\n}.
          qq{</component>\n};
    }
    else {
        croak "E040: Panic -- Invalid type ('$type'), expected ('prog' or 'func')";
    }

    my $file_name = 'S-'.$dat_sha1.'.txt';
    my $file_full = $VBRepo.'\\'.$file_name;

    unless (-f $file_full) {
        write_file($file_full, $file_content);
    }

    if ($type eq 'func') {
        my $obj = Win32::OLE->GetObject('script:'.$file_full) or croak "E050: ",
          "Couldn't Win32::OLE->GetObject('script:$file_full')",
          " -> ".Win32::GetLastError;

        for (keys %dat_func) {
            $dat_func{$_} = sub { $obj->$_(@_); };
        }
    }

    bless {
      'name' => $file_name,
      'type' => $type,
      'lang' => $lang,
      'func' => \%dat_func,
    }, $pkg;
}

sub compile_prog_vbs {
    my ($code) = @_;
    Win32::VBScript->new('prog', 'vbs', $code);
}

sub compile_prog_js {
    my ($code) = @_;
    Win32::VBScript->new('prog', 'js', $code);
}

sub compile_func_vbs {
    my ($code) = @_;
    Win32::VBScript->new('func', 'vbs', $code);
}

sub compile_func_js {
    my ($code) = @_;
    Win32::VBScript->new('func', 'js', $code);
}

sub _run {
    my $self = shift;
    my ($scr) = @_;

    unless ($scr eq 'cscript' or $scr eq 'wscript') {
        croak "E060: Invalid script ('$scr'), expected ('cscript' or 'wscript')";
    }

    my $fn   = $self->{'fn'};
    my $full = $VBRepo.'\\'.$fn;
    my $lang = $self->{'lang'};

    unless (-f $full) {
        croak "E070: Panic -- can't find executable '$full'";
    }

    my $engine =
      $lang eq 'vbs' ? 'VBScript' :
      $lang eq 'js'  ? 'JScript'  :
      croak "E080: Panic -- invalid language ('$lang'), expected ('vbs' or 'js')";

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

sub func {
    my $self  = shift;
    my $mname = shift;

    $self->{'func'}{$mname};
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