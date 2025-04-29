# Naming Convention
The naming convention is as follows:

Starting with a file name with extension, but without path to it in its name:

- The file name is split by each dot. 
- The first part is the display name.
- Each following part (except the last for files, i.e. non-folders) is either:
-  - a MIME type (if it's uppercase)
-  - a file extension (if it's lowercase)
- The last part (if it is a file extension) is ignored.



## File Extension Schema
For example, "Python_Note.py.txt" is a text file that will only be shown if only *.py files 
are selected, or if no selection is made, if all files within the folder have the extension .py.

## MIME Type Schema
As for the MIME-types, the expected format is "<rough type>", or "<rough type>..<subtype>".<br>
The understood MIME-types are pulled from https://github.com/samuelneff/MimeTypeMap/blob/master/MimeTypeMap.cs.<br>
If a MIME type isn't understood, the output will be "application/octet-stream" (where for example application is the rough type, and octet stream the subtype).

### Examples

A simple example would be "convert_to_mp3.AUDIO.py". This will only show the context menu "convert_to_mp3" if only audio files are selected (or, if no selection is made, if all files in the folder are audio files).

Similarly, "convert_to_mp3.AUDIO.VIDEO.py" will only show for audio or video files.

And "convert_to_mp3.AUDIO..OGG.py" will only show if the MIME-type is "audio/ogg", which is e.g. the case for .opus, but not for .mp3 (which is audio/mpeg).


# Tips


If you want to call a Python script,
it should be .py if the script should run in the foreground, or .pyw if the script should run in the background.
Each middle part should have all its letters in lower case if it's a file extension, and in upper case if it's a
MIME type.

