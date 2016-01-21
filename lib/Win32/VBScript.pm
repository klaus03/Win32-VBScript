package Win32::VBScript;

use strict;
use warnings;

use Carp;
use Digest::SHA qw(sha1_hex);
use File::Slurp;
use Win32::OLE;

require Exporter;
our @ISA         = qw(Exporter);
our %EXPORT_TAGS = ('ini' => [qw(
    compile_prog_vbs compile_prog_js
    compile_func_vbs compile_func_js
)]);
our @EXPORT      = qw();
our @EXPORT_OK   = ( @{ $EXPORT_TAGS{'ini'} } );

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
    my $dat_class = "InlineWin32COM.WSC\\_$dat_sha1.wsc";

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
              qq{description="Inline::WSC Class" }.
              qq{progid="$dat_class" }.
              qq{version="1.0">\n}.
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

    my $file_name = 'T_'.$dat_sha1.'.txt';
    my $file_full = $VBRepo.'\\'.$file_name;

    write_file($file_full, $file_content);

    if ($type eq 'func') {
        my $obj = Win32::OLE->GetObject('script:'.$file_full) or croak "E050: ",
          "Couldn't Win32::OLE->GetObject('script:$file_full')",
          " -> ".Win32::GetLastError;

        for my $method (keys %dat_func) {
            $dat_func{$method} = sub { $obj->$method(@_); };
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
    my ($scr, $mode) = @_;

    unless ($scr eq 'cscript' or $scr eq 'wscript') {
        croak "E060: Invalid script ('$scr'), expected ('cscript' or 'wscript')";
    }

    my $name = $self->{'name'};
    my $lang = $self->{'lang'};
    my $type = $self->{'type'};

    unless ($type eq 'prog') {
        croak "E065: Invalid type ('$type'), expected ('prog')";
    }

    my $full = $VBRepo.'\\'.$name;

    unless (-f $full) {
        croak "E070: Panic -- can't find executable '$full'";
    }

    my $engine =
      $lang eq 'vbs' ? 'VBScript' :
      $lang eq 'js'  ? 'JScript'  :
      croak "E080: Panic -- invalid language ('$lang'), expected ('vbs' or 'js')";

    my @param = ($scr, '//Nologo', '//E:'.$engine, $full);

    if ($mode eq 'a') {
        system 1, @param; # asynchronous
    }
    elsif ($mode eq 's') {
        system    @param; # sequentially
    }
    else {
      croak "E082: Panic -- invalid mode ('$mode'), expected ('a' or 's')";
    }
}

sub cscript {
    my $self = shift;
    $self->_run('cscript', 's'); # s = sequentially
}

sub wscript {
    my $self = shift;
    $self->_run('wscript', 's'); # s = sequentially
}

sub async {
    my $self = shift;
    $self->_run('wscript', 'a'); # a = asynchronous
}

sub func {
    my $self  = shift;
    my $mname = shift;

    $self->{'func'}{$mname};
}

sub flist {
    my $self  = shift;
    my $sf = $self->{'func'};

    sort grep { $sf->{$_} } keys %$sf;
}

1;

__END__

=head1 NAME

Win32::VBScript - Run Visual Basic programs

=head1 DESCRIPTION

This module allows you to invoke code fragments written in Visual
Basic (or even JavaScript) from within a perl program.
The Win32::OLE part has been copied from Inline::WSC.

=head1 SYNOPSIS

    use strict;
    use warnings;

    use Win32::VBScript qw(:all);

    # This is the procedural interface:
    # *********************************

    my $p1 = compile_prog_js ([ qq{WScript.StdOut.WriteLine("Bonjour");} ]); cscript($p1);
    my $p2 = compile_prog_vbs([ qq{WScript.StdOut.WriteLine "Hello"}     ]); cscript($p2);

    # This is the OO interface:
    # *************************

    compile_prog_js ([ qq{WScript.StdOut.WriteLine("Test1");} ])->cscript;
    compile_prog_vbs([ qq{WScript.StdOut.WriteLine "Test2"}   ])->cscript;

    # And with wscript, of course, you can use MsgBox:
    # ************************************************

    compile_prog_vbs([ qq{MsgBox "Test3"} ])->wscript;

    # You can even define functions in Visual Basic...
    # ************************************************

    my $t1 = compile_func_vbs([ <<'EOF' ]);
      ' Say hello:
      Function Hello(ByVal Name)
        Hello = ">> " & Name & " <<"
      End Function

      ' Handy method here:
      Function AsCurrency(ByVal Amount)
        AsCurrency = FormatCurrency(Amount)
      End Function
    EOF

    # ...or even JavaScript...
    # ************************

    my $t2 = compile_func_js([ <<'EOF' ]);
      function greet(name) {
        return "Greetings, " + name + "!";
      } // end greet(name)
    EOF

    # ...and call the functions later in Perl:
    # ****************************************

    print 'Compiled functions are: (',
      join(', ', map { "'$_'" }
      sort { lc($a) cmp lc($b) } $t1->flist, $t2->flist),
      ')', "\n\n";

    {
        no strict 'refs';

        *{'::hi'}  = $t1->func('Hello');
        *{'::cur'} = $t1->func('AsCurrency');
        *{'::grt'} = $t2->func('greet');
    }

    print hi('John'), ' gets ', cur(100000), ' -> ', grt('Earthling'), "\n\n";

=head1 AUTHOR

Klaus Eichner <klaus03@gmail.com>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2016 by Klaus Eichner

All rights reserved. This program is free software; you can redistribute
it and/or modify it under the terms of the artistic license 2.0,
see http://www.opensource.org/licenses/artistic-license-2.0.php

=cut
