# extract_and_install_fonts
python script to automate the task of extracting zip files and then install font files and then delete the downloaded files.
For a long time, I was stuck doing the same repetitive and boring task as a part-time graphic designer: downloading font files, manually extracting each ZIP, installing every single TTF or OTF file, and then cleaning up the leftover files.
It sounds simple, but doing this over and over again becomes a real chore.
So, I decided to automate the entire process with a Python script. Here’s how it works:
This is the script, and here are the ZIP files I’ve downloaded. Now, all I have to do is run the script, paste the folder path where the ZIP files are stored, and let it do the rest.
The script automatically:
•
Searches for the ZIP files in the specified folder
•
Extracts them
•
Installs any TTF or OTF fonts that aren’t already in the system
•
Cleans up by deleting the downloaded files
•
And finally, refreshes the system’s font cache
To prove it works, let’s check Adobe Illustrator. As you can see, all the fonts are installed—no manual refreshing needed.
This simple script has saved me so much time, and I hope it can help others too.
If you’d like the full code, feel free to comment or DM me.
Thanks for watching!

1. copy the code to any IDE
2. run as administrator
3. copy and enter the file path
4. and all set!
