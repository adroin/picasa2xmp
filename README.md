# picasa2xmp
Picasa2xmp is a VB Script that converts .picasa.ini information about rating and albums 
into XMP format (AfterShot Pro 3).

How it works from a functional point of view :

Picasa stores its data inside .picasa.ini files that are located in each picture folders.
Those files contain information about the albums themselves, but also about the pictures.
They tell all the functionnalities that have been applied to each picture inside Picasa
like autolevel, fill light, etc...
picacasa2xmp only works with two kind of information : the rating (whether the picture has
a star inside Picasa or not) and the Picasa album(s) the picture belongs to. 
Picasa2xmp will transalte the rating information as follow : each Picasa Star will give a 
rating of 3 stars inside AfterShot Pro.
Picasa2xmp will also migrate the albums into Keywords inside AfterShot Pro. For this to
work properly, you'll have to fill a text file that gives the corresponding XMP Keywords
for each of your Picasa albums. See the "procedure.txt" file inside this repository for more
information on how to proceed.

How it works from a technical point of view :

Picasa2xmp won't generate the XMP file itself, so you'll have to generate it for the entire
catalog inside AfterShot Pro. You'll probably want to create a dedicated catalog that matchs the
list of folders you want to convert.
Once you have your AfterShot XMP files ready, just launch the VB Script and choose the main
folder to convert. Picasa2xmp will rewrite XMP files when it discovers albums attachments or
stars inside the .picasa.ini file. The "procedure.txt" file inside this project explains 
what are the technical prerequisites for the script to work properly.
