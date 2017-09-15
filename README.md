# LookupAccountName-NoArgs
LookupAccountNameA from advapi32.dll - as an easy to use application include VB6 source and ready to use binary (along with Windows 10 compatible manifest)

but it uses <code>GetUserName</code> (<code>GetUserNameA</code> from <code>advapi32.dll</code>) and <br/>
using <code>GetComputerName</code> (<code>GetComputerNameA</code> from <code>kernel32.dll</code>) internally,<br/>
so it won't accepts any arguments.

<hr/>

this is a simpler usage of <code>LookupAccountNameA</code>!

<hr/>

you should look in this repository: <a href="http://github.com/eladkarako/LookupAccountName-WithArgs/">LookupAccountName-WithArgs</a>,<br/>
it accepts both machine-name which can be empty-string and user-name (must!).<br/>

and after a lookup that make include a network access it will return both the sid string and the domain the user-name was found in (comma separated),
which is more close to the native use of <code>LookupAccountNameA</code>.