function onOpen() {
 DocumentApp
   .getUi()
   .createMenu("Convert Document")
   .addItem("Convert to HTML", "convertGoogleDocToCleanHtml")
   .addItem("Settings", "showSettings")
   .addToUi();
}

function showSettings(){
  var html = HtmlService.createTemplateFromFile('dialog').evaluate();
  var htmlOutput = HtmlService
    .createHtmlOutput(html.getContent())
    .setWidth(350)
    .setHeight(350);
  DocumentApp.getUi().showModalDialog(htmlOutput, 'Settings');
}

function storeSettings(settings){
  let defaults = defaultSettings();
  try {
    const properties = PropertiesService.getScriptProperties();
    for (const key in defaults) {
      let v = settings[key] ? true : false;
      properties.setProperty(key, v);
    }    
  } catch (err) {
    // TODO (developer) - Handle exception
    console.log("Failed with error " + err.message);
  }
}

function defaultSettings(){
  return {
    'ExportFonts': {
      value: false,
      label: "Include font information in converted HTML"
    },

    'CustomizeTable': {
      value: false,
      label: "Add column width information to tables"
    },

    'ItalicParagraphsAsBlockquotes': {
      value: false,
      label: "Output paragraphs in italics as blockquote"     
    },

    'OutputTextInMonospacedAsCode': {
      value: false,
      label: "Output text using Monospaced fonts as code"     
    },

    'ExportImages': {
      value: false,
      label: "Export inline images to Google Drive folder named 'Exported Images'"     
    }      

  } 
}

function retrieveSettings(){
  let settings = defaultSettings();
  try {
    const properties = PropertiesService.getScriptProperties();
    const data = properties.getProperties();
    for (const key in data) {
      if (settings[key]){
        settings[key].value = data[key] == "true" ? true : false;
      }
    }
  } catch (err) {
    // TODO (developer) - Handle exception
    console.log('Failed with error %s', err.message);
  }
  return settings;
}


function showConversionInSidebar(html) {
  var widget = HtmlService.createHtmlOutput("<h1>Converted Document</h1><textarea style=\"width: 95%;height: 50em\">"+html+"</textarea>"+html);
  DocumentApp.getUi().showSidebar(widget);
}

function convertGoogleDocToCleanHtml() {
  var body = DocumentApp.getActiveDocument().getBody();
  var numChildren = body.getNumChildren();
  var output = [];

  // Walk through all the child elements of the body.
  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    output.push(processItem(child));
  }

  var html = output.join('\n');

  showConversionInSidebar(html);
}

function processItem(item) {
  var settings = retrieveSettings();
  var output = [];
  var prefix = "",
    suffix = "",
    styles = "";

  //paragraph
  if (item.getType() == DocumentApp.ElementType.PARAGRAPH) {
    switch (item.getHeading()) {
      // Add a # for each heading level. No break, so we accumulate the right number.
      case DocumentApp.ParagraphHeading.HEADING6:
        prefix = "<h6>", suffix = "</h6>";
        break;
      case DocumentApp.ParagraphHeading.HEADING5:
        prefix = "<h5>", suffix = "</h5>";
        break;
      case DocumentApp.ParagraphHeading.HEADING4:
        prefix = "<h4>", suffix = "</h4>";
        break;
      case DocumentApp.ParagraphHeading.HEADING3:
        prefix = "<h3>", suffix = "</h3>";
        break;
      case DocumentApp.ParagraphHeading.HEADING2:
        prefix = "<h2>", suffix = "</h2>";
        break;
      case DocumentApp.ParagraphHeading.HEADING1:
        prefix = "<h1>", suffix = "</h1>";
        break;
      default:
        prefix = "<p>", suffix = "</p>";
    }

    if (item.getNumChildren() == 0)
      return "";
  
  //images
  } else if (item.getType() == DocumentApp.ElementType.INLINE_IMAGE) {
    processImage(item, output, settings);

  //Tables
  } else if (item.getType() === DocumentApp.ElementType.TABLE) {
    if (!settings['CustomizeTable'].value){
      prefix = "<table>", suffix = "</table>";
    } else {
      prefix = "<table class=\"table-bordered\" style=\"width: 100%;\">", suffix = "</table>";
    }
  } else if (item.getType() === DocumentApp.ElementType.TABLE_ROW){
    prefix = "<tr>", suffix = "</tr>";
  } else if (item.getType() === DocumentApp.ElementType.TABLE_CELL){
    let styles = '';
    if (settings['CustomizeTable'].value){
      switch (item.getParentRow().getChildIndex(item)) {
        case 0:
          styles = " style=\"width: 12%;\"";
          break;
        case 1:
          styles = " style=\"width: 88%;\"";
          break;
        default:
          styles = ""
      }  
    }
    prefix = "<td"+styles+">", suffix = "</td>";
  
  //Lists
  } else if (item.getType() === DocumentApp.ElementType.LIST_ITEM) {
    //find list type
    var listType = [DocumentApp.GlyphType.BULLET, 
      DocumentApp.GlyphType.HOLLOW_BULLET, 
      DocumentApp.GlyphType.SQUARE_BULLET].includes(item.getGlyphType());

    // First list item
    if (!item.getPreviousSibling() || item.getPreviousSibling().getType() != DocumentApp.ElementType.LIST_ITEM || item.getPreviousSibling().getNestingLevel() < item.getNestingLevel() ){ 
      prefix = (listType ? '<ul>' : '<ol>') +'<li>', suffix = "</li>";
    // Last main list item
    } else if (item.isAtDocumentEnd() || !item.getNextSibling() || item.getNextSibling().getType() != DocumentApp.ElementType.LIST_ITEM) {
      prefix = "<li>", suffix = "</li>"+(listType ? '</ul>' : '</ol>');
    // Last sub list item
    } else if (item.getNextSibling() && item.getNextSibling().getType() == DocumentApp.ElementType.LIST_ITEM && item.getNextSibling().getNestingLevel() < item.getNestingLevel()) {
      prefix = "<li>", suffix = "</li>"+(listType ? '</ul>' : '</ol>'); 
    // normal item
    } else {
      prefix = "<li>", suffix = "</li>";
    }
  }

  output.push(prefix);

  if (item.getType() == DocumentApp.ElementType.TEXT) {
    processText(item, output, settings);
  } else {

    if (item.getNumChildren) {
      var numChildren = item.getNumChildren();

      // Walk through all the child elements of the item.
      for (var i = 0; i < numChildren; i++) {
        var child = item.getChild(i);
        output.push(processItem(child));
      }
    }
  }

  output.push(suffix);
  return output.join('');
}


function processText(item, output, settings) {
  var text = item.getText();
  var indices = item.getTextAttributeIndices();

  if (indices.length <= 1) {
    // Assuming that a whole para fully italic is a quote
    if (item.isBold()) {
      output.push('<strong>' + text + '</strong>');
    } else if (item.isItalic() && settings['ItalicParagraphsAsBlockquotes'].value) {
      output.push('<blockquote>' + text + '</blockquote>');
    } else if (item.isItalic()) {
      output.push('<i>' + text + '</i>');
    } else if (text.trim().indexOf('http://') == 0) {
      output.push('<a href="' + text + '" rel="nofollow">' + text + '</a>');
    } else if (text.trim().indexOf('https://') == 0) {
      output.push('<a href="' + text + '" rel="nofollow">' + text + '</a>');
    } else {
      output.push(text);
    }
  } else {

    for (var i = 0; i < indices.length; i++) {  
      var partAtts = item.getAttributes(indices[i]);    
      var startPos = indices[i];
      var endPos = i + 1 < indices.length ? indices[i + 1] : text.length;
      var partText = text.substring(startPos, endPos);      

      if (partAtts.ITALIC) {
        output.push('<i>');
      }
      if (partAtts.BOLD) {
        output.push('<strong>');
      }
      if (partAtts.UNDERLINE) {
        output.push('<u>');
      }

      //post processing of text snippets
      if (partAtts.FONT_FAMILY && partAtts.FONT_FAMILY.match(/courier|mono/i) && settings['OutputTextInMonospacedAsCode'].value) {
        output.push('<code>' + partText + '</code>');
      } else if (partAtts.FONT_FAMILY && settings['ExportFonts'].value) {
        output.push('<span style="font-family: '+partAtts.FONT_FAMILY+'">' + partText + '</span>');
      } else if (partText.trim().indexOf('http://') == 0) {
        output.push('<a href="' + partText + '" rel="nofollow">' + partText + '</a>');
      } else if (partText.trim().indexOf('https://') == 0) {
        output.push('<a href="' + partText + '" rel="nofollow">' + partText + '</a>');
      } else {
        output.push(partText);
      }

      if (partAtts.ITALIC) {
        output.push('</i>');
      }
      if (partAtts.BOLD) {
        output.push('</strong>');
      }
      if (partAtts.UNDERLINE) {
        output.push('</u>');
      }

    }
  }
}

function processImage(item, output, settings) {
  var blob = item.getBlob();
  var contentType = blob.getContentType();
  var extension = "";
  if (/\/png$/.test(contentType)) {
    extension = ".png";
  } else if (/\/gif$/.test(contentType)) {
    extension = ".gif";
  } else if (/\/jpe?g$/.test(contentType)) {
    extension = ".jpg";
  } else {
    throw "Unsupported image type: " + contentType;
  }
  
  var b64Url='data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());

  var image_info = ''
  if (settings['ExportImages'].value){
    var file = saveImageToDrive(blob, extension);
    image_info = ' data-gid="'+file.getId()+'" data-file-name="'+file.getName()+'"';
  }

  output.push('<img'+image_info+' style="width: '+item.getWidth()+'px" src="'+b64Url+'" />');
}

function saveImageToDrive(image, extension) {
  var folderName = "Exported Images";
  var folderIterator = DriveApp.getRootFolder().getFoldersByName(folderName);
  var folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.getRootFolder().createFolder(folderName);

  var name = DocumentApp.getActiveDocument().getName().toLowerCase().replace(/[^\w]/g, '-')+"_"+(Math.random() * Date.now()).toString(36)+extension;
  image.setName(name);
  var file = folder.createFile(image);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file;
}