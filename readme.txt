A script to send mails - obviously.

Its got a two-fold purpose.

If its run without parameters, no code is executed, and it makes it easy to . dot source
the functions in the script, mainly the Send-Mail advanced function, which makes up the
core of the script.

If parameters is supplied, those will be passed on to the Send-Mail function, and as
such the script can also be used directly from the command line.

Constructed as Advanced functions all the way through, so common parameters is available,
although the -WhatIf wont do anything.