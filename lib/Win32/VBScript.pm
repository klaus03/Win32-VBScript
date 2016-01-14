package Win32::VBScript;

use strict;
use warnings;

require Exporter;
our @ISA         = qw(Exporter);
our %EXPORT_TAGS = ('all' => [qw()]);
our @EXPORT      = qw();
our @EXPORT_OK   = ( @{ $EXPORT_TAGS{'all'} } );

sub new {
    my $pkg = shift;

    bless {}, $pkg;
}

1;

__END__

=head1 NAME

Win32::VBScript - Run Visual Basic programs

=head1 SYNOPSIS

    use Win32::VBScript;

    my $prog = Win32::VBScript->new('vbs', [ qq{MsgBox "Greetings Earthlings..."} ]);
    $prog->cscript;

=head1 DESCRIPTION

This module allows you to invoke Visual Basic Programs from within a perl program.

=head1 AUTHOR

Klaus Eichner <klaus03@gmail.com>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2016 by Klaus Eichner

All rights reserved. This program is free software; you can redistribute
it and/or modify it under the terms of the artistic license 2.0,
see http://www.opensource.org/licenses/artistic-license-2.0.php

=cut
