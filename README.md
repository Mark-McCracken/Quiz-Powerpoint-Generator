# Quiz Powerpoint Generator

Making this quiz powerpoint takes a lifetime!
About 10 hours the first time you come round to doing it, and it's fraght with errors,
as soon as you realise you need to make a change, you need to change it in multiple places.
So much formatting and messing around with animations when you make a change, changing font sizes, making sure question numbers and answer numbers are lined up.

I hate all of that, I just want to come up with the questions and do the presenting!

So I made this script to automate every last thing that I could about making the presentation.

Questions and answers all go in an excel workbook, super straightforward to understand.

Picture round, just stick the pictures (must be .jpg/.jpeg/.png) in a folder and name them with the answer

Music round, create a folder with the name of the song and artist.
I then get my mp3/m4a file, open it in quicktime editor, trim it down to the right clip for the question, and export as question.m4a into the foler, then trim again for the answer, and export as answer.m4a into the same folder (although this is optional, if there's no answer.m4a, it'll just use the question snippet again on the answer slide.)
I slice the answer track so the most obvious part of the song plays at around seconds, which is when the 2nd image animates in, so the answer becomes evident at the same time as the picture.
I just google 2 images add them as img_1.jpep (or .png) and img_2.jpeg to the same folder.

## Quiz question tips
Make sure you have at least a few questions in each worded round that are multiple choice, and a few "list 5 of these" type questions, to give chances for half points.
Try to make sure the music round isn't all the same genre, and has a good mix of men/women, old/new, pop, classical, hip-hop, rock, dance, yada ya. Variety is good here.
And avoid questions that might be potentially offensive.

## Running
You need to install the python library python-pptx.
Then run python fill_template.py, and output.pptx should be generated.

## What do I need to do after that?

### Autosizing text
The text should autosize, but because of the way the template gets populated, they overflow at first.

To fix, click on the text, make a tiny change and it should auto-resize.

### Check images
Sometimes the images for picture round aren't very nicely laid out, they automatically scale to fit and crop the edges.
So long as you pick images that are close to the same dimension ratio, should be mostly fine, but might want to tweak those to be sure it's clear.

### Check any currency answers
Currency isn't stored in the excel cell value, so if you enter £4, it might be shown on the slide as just 4, so double check those just to be sure.

### Animating the sound
The sound in each music round question and answer slide does not start
(the python library does not support animations, or media placeholders)

To fix, select each media item on each slide, and have it start "with previous".

### Answer sheets
I then need to print out some answer sheets, modifying the round names. But this is a tiny job I just do manually.

### Bonus Point
Last point, I usually have 59 questions.
The last point can only be won by one team - it's the first team to spot something hidden in the slides.
I normally take a picture of Nicolas Cage's head, and photoshop it into one of the answer round slides - first team to stand up and shout "Nicolas Cage!" wins the point.
Don't forget to mention this in the rules! But doesn't have to be Nic Cage, could be anything, or can leave it out.
