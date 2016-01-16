use strict;
use warnings;

use Test::More tests => 7;

use Capture::Tiny qw(capture_merged);

use_ok('Win32::VBScript', ':all');

{
    my $txt = capture_merged {
      compile_prog_js([ qq{WScript.StdOut.WriteLine("Hi JS");} ])->cscript;
    };

    $txt =~ s{\n}''xmsg;
    is($txt, 'Hi JS', 'compile_prog_js() works');
}

{
    my $txt = capture_merged {
      compile_prog_vbs ([ qq{WScript.StdOut.WriteLine "Hi VBS"} ])->cscript;
    };

    $txt =~ s{\n}''xmsg;
    is($txt, 'Hi VBS', 'compile_prog_vbs() works');
}

{
    my $comp = compile_func_js([ q~
      function Greet(name) {
        return "Greetings, " + name + "...";
      } // end Greet(name)~
    ]);

    my $lst = join(', ', $comp->flist);
    is($lst, 'Greet', 'flist (JS) works');

    my $txt = $comp->func('Greet')->('Earthling');
    is($txt, 'Greetings, Earthling...', 'comp->func (JS) works');
}

{
    my $comp = compile_func_vbs([ q~
      ' Say hello:
      Function Hey(ByVal Name)
        Hey = "Hey " & Name & " !!"
      End Function~
    ]);

    my $lst = join(', ', $comp->flist);
    is($lst, 'Hey', 'flist (Vbs) works');

    my $txt = $comp->func('Hey')->('Jude');
    is($txt, 'Hey Jude !!', 'comp->func (Vbs) works');
}
