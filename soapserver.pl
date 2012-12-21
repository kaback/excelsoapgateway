#!perl -w

# found here:
# http://guide.soaplite.com/#more%20complex%20server%20%28daemon,%20mod_perl%20and%20mod_soap%29
#
use SOAP::Transport::HTTP;

use ExcelSoapGateway;

# don't want to die on 'Broken pipe' or Ctrl-C
$SIG{PIPE} = $SIG{INT} = 'IGNORE';

my $daemon = SOAP::Transport::HTTP::Daemon
   -> new (LocalPort => 82)
   -> dispatch_to('ExcelSoapGateway')
;

print "Contact to SOAP server at ", $daemon->url, "\n";
$daemon->handle;


