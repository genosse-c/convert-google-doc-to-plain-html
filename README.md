# convert-google-doc-to-plain-html
Google app script to convert a google document to plain html

## Usage
1. Create a new google doc in your drive.
2. Use Extensions -> App Script to create a new container bound app script for the google doc
3. Create a Code.gs file and copy the content of the Code.js file into it
4. Create a dialog.html file and copy the content of dialog.html into it.
5. Select the Code.gs file and run any of the functions in it using the menu at the top of the editor. 
   This will bring up a dialog asking permission for the app script to access your google drive.
6. Now copy the content of any google doc you need as plain HTML into the document und use the *Convert Document* menu to create HTML output

This will open a sidebar with a text area containing the HTML, and below it a rendered version of that HTML. 
Use the contents of the textarea if you need the raw HTML, use the rendered HTML if you intend to paste into a rich text editor

You can set some options using "Setting" in the *Convert Document* menu.

### Images
Images that are part of the source document can be exported as well. The script will create a folder *Exported Images* 
into which all the images in the document are saved (depending on a setting).
If you require the images to work in your exported HTML, you will need to store them on your web server 
and adjust the path to the images manually.

## The Usual
This code comes with no guarantee, but has worked for me. Some of the options create HTML that cater to my specific needs.
You might need to change some of the code in the script to adopt it to yours.
