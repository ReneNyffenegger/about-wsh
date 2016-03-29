option explicit

dim args
dim nofArguments
dim arg

set args         = wscript.arguments
nofArguments     = args.count

wscript.echo ("The script was passed " & nofArguments & " arguments")

for arg = 1 to nofArguments 

   wscript.echo ("Argument " & arg & ": " & args(arg-1))

next
