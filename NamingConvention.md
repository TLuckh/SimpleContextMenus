# Naming Convention
Start with a file representing a context menu, or a folder representing a context drop-down menu. 
Let its name be given with extension, but without path to it in its name.

Then the naming convention is as follows:
- The name is split by each dot.
- The first part is the display name.
- Each following part (except the last for files, i.e. non-folders) is either:
    - a MIME type (if it's uppercase)
    - a file extension (if it's lowercase)
- The last part (if it is a file extension) is ignored.

## Matching Criteria
If no type extensions or MIME types were given, the context menu will always be shown.
Otherwise, the context menu will only be shown if of all the selected files (or, if no selection was made, all files in the folder),
at least one matches any of the given MIME types or file extensions.

## File Extension Schema

For example, "Python_Note.py.txt" is a text file that will only be shown if at least one *.py file
is selected, or if no selection is made, if any file within the folder has the extension .py.

## MIME Type Schema

As for the MIME-types, the expected format is "<rough type>", or "<rough type>..<subtype>".<br>
The understood MIME-types are pulled from https://github.com/samuelneff/MimeTypeMap/blob/master/MimeTypeMap.cs.<br>
If a MIME type isn't understood, the output will be "application/octet-stream" (where for example application is the
rough type, and octet stream the subtype).

### Examples

A simple example would be "convert_to_mp3.AUDIO.py". This will only show the context menu "convert_to_mp3" if at least one audio
file is selected (or, if no selection is made, if any file in the folder is an audio files).

Similarly, "convert_to_mp3.AUDIO.VIDEO.py" will only show if at least one audio or video file is selected.

And "convert_to_mp3.AUDIO..OGG.py" will only show if the MIME-type is "audio/ogg", which is e.g. the case for .opus, but
not for .mp3 (which is audio/mpeg).

# Tips

If you want to call a Python script,
it should be .py if the script should run in the foreground, or .pyw if the script should run in the background.
Each middle part should have all its letters in lower case if it's a file extension, and in upper case if it's a
MIME type.


