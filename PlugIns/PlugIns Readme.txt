To Install a Plug-in DLL:

COMPILE TO DLL THEN...
1. Copy the DLL to Axiom\PlugIns folder
2. Register the DLL using RegSvr32,At a command prompt type:
	REGSVR32 full-path-to-dll-file


The DLL should adhere to the following:

1. The DLL can have only one Class, if more exist only the first will be used.
2. By default every Function is supposed to have only one Argument, if a function has more than one this should be stated in the procedure attribute.

for example:

Apply Font Attributes to Text
Inputs:Text,Font Name,Font Color,FontSize

The first line is the function description, the second is the arguments discripion.

