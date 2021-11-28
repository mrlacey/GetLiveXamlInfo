# Get Live XAML Info

This was an experiment and is no longer available (as a preview) in the marketplace.

It was created as a way to extract what Visual Studio thought was the equivalent XAML of a rendered UI. I initially used this as a way to compare what I thought should be generated versus what was actually created by a tool.

It was useful to an extent but due to the way it used reflection to access non-public members (the only way it was able to access the needed information) it was very fragile and would need code changes with (what felt like) every minor release of Visual Studio, it was impractical to support and maintain. 

This code remains for reference but I don't expect to ever do anything with this again.

----

It does things that are unsupported, it uses things that are undocumented, and accesses private assemblies and internal code that it shouldn't.

No promises or guarantees about anything to do with this code or its use.

It works on my machine and that's all I need. Sorry.
