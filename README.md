# Quiz Powerpoint Generator

Making this quiz powerpoint takes a lifetime!
About 10 hours the first time you come round to doing it, and it's fraght with errors,
as soon as you realise you need to make a change, you need to change it in multiple places.
So much formatting and messing around with animations when you make a change, changing font sizes, making sure question numbers and answer numbers are lined up.

I hate all of that, I just want to come up with the questions and do the presenting!

So I made this script to automate every last thing that I could about making the presentation.

Questions and answers all go in an excel workbook, super straightforward to understand.

Picture round, just stick the pictures (must be .jpg/.jpeg/.png) in a folder and name them with the answer

Music round, create a folder with the name of the song and artist, and add the question track (must be .mp3/.m4a), and optionally a different answer track if you want people to hear the chorus, and 2 images for the answer slide.

Hit run, and it should spit out the output file.

## What do I need to do after that?

### Autosizing text
The text should autosize, but because of the way the template gets populated, they overflow at first.

To fix, click on the text, make a tiny change and it should auto-resize.

### Animating the sound
The sound in each music round question and answer slide does not start (python library does not support animations, or media placeholders.)

To fix, select each media item on each slide, and have it start "with previous".
