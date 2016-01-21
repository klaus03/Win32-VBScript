use strict;
use warnings;

use Test::More tests => 9;

use Capture::Tiny qw(capture_merged);

use_ok('Win32::VBScript', qw(:ini));

{
    my $txt = capture_merged {
      compile_prog_js([ qq{WScript.StdOut.WriteLine("Hi JS pl_cscript");} ])->pl_cscript;
    };

    $txt =~ s{\n}''xmsg;
    is($txt, 'Hi JS pl_cscript', 'compile_prog_js(PL) works');
}

{
    my $txt = capture_merged {
      compile_prog_vbs ([ qq{WScript.StdOut.WriteLine "Hi VBS pl_cscript"} ])->pl_cscript;
    };

    $txt =~ s{\n}''xmsg;
    is($txt, 'Hi VBS pl_cscript', 'compile_prog_vbs(PL) works');
}

{
    my $txt = capture_merged {
      compile_prog_js([ qq{WScript.StdOut.WriteLine("Hi JS ms_cscript");} ])->ms_cscript;
    };

    $txt =~ s{\n}''xmsg;
    is($txt, 'Hi JS ms_cscript', 'compile_prog_js(MS) works');
}

{
    my $txt = capture_merged {
      compile_prog_vbs ([ qq{WScript.StdOut.WriteLine "Hi VBS ms_cscript"} ])->ms_cscript;
    };

    $txt =~ s{\n}''xmsg;
    is($txt, 'Hi VBS ms_cscript', 'compile_prog_vbs(MS) works');
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
