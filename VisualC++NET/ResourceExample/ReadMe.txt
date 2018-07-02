Resource Example
================

(c) Richard Grimes 2003

This example illustrates how to get round some of the problems using
managed resources with C++. 

Non-Form Resources
------------------
You can add either a compiled resource or a non compiled resource.
A compiled resource has the advantage that it can be used with 
ResourceManager to extract a specific resource item. To do this, use
the Add New Item context menu to add a Assembly Resource File (.resx). In this example I have added a 
resource called StringData which has various named strings. The wizard
generated project settings will add this resource to the assembly with
the name ResourceExample.ResourceFiles.resources. It assumes that this
will be the only managed resource for the assembly (I will return to
this later) and it has that strange string "ResourceFiles". 

Open the property pages for the StringData.resx file and on the Managed 
Resources page edit the Resource File Name property for all 
configurations to be:

$(IntDir)/$(RootNamespace).resources

This has now renamed the managed resource to be ResourceExample.resources
and so the name used in the ResourceManager constructor is simply 
ResourceExample.

Multiple Embedded Resources
---------------------------
As indicated above the wizard will give the resource name based on the name
of the assembly. If you want to add another compiled resource then you
can simply edit the Resource File Name property for all configurations and
give the compiled resource file another name. This is what I have done with
the file OtherResources.resx.

Linked Resources
----------------
resx files added to the project will always be added as an embedded resource.
If you want to add the resource as a linked resource then you have to turn
off compilation of the resx file and add a separate pre-link build and then
add the resource through the linker command line. Here are the steps to add
the LinkedResource compiled resource:

1) Use the Add New Item context menu to add a Assembly Resource File (.resx)

2) On the General property page for the LinkedResource.resx file change the
Excluded from Build to Yes for all configurations.

3) On the project's property pages for all configurations add a custom build 
step. Go to Build Events and on the Pre-Link Event add the following for the
Command Line : 

resgen LinkedResource.resx "$(OutDir)\LinkedResource.resources"

if you have more than one linbed compiled resource use the /compile switch 
and provide pairs of strings with the input name and output name

4) On the Linker Command Line page for all configurations add 

/assemblylinkresource:"$(OutDir)\LinkedResource.resources"





