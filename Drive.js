/**
 * Settings
 */
//array of objects each with 
// a *regex* property to find up to three parts of name metadata in the filename
// a *fields* property consisting of an array naming the fields: "title", "author", "year"
const metadata_filter = [{
    regex: /^(.+)\[([^\)]+)\].*\((\d+)\)/ms, //<title>[<author>](<year>)
    fields: ['title', 'author', 'year']
  },  
   {
    regex: /^(.+)[\s_\-]by[\s_\-](.+)/ms, //<title>by<author>
    fields: ['title', 'author']
  },
  {
    regex: /^(.+)[_]*\[([^\]]+)\]/ms, //<title>[<author>]
    fields: ['title', 'author']
  },
  {
    regex: /^\[(.+)\][\s\-_]*(.+)/ms, //[<author>]<title>
    fields: ['author', 'title']
  },
  {
    regex: /^(.+)[\s\-_]*\(([^\)]+)\)/ms, //<title>(<author>)
    fields: ['title', 'author']
  }  
];

//Tests whether the book name follows the "{title}_[{author}]_({year})" format with year being optional
const standard_filter = /^([^\[]+)[_]*\[([^\]]+)]/m 

const default_settings = {
  'special_folder_name': {
    value: 'Zu verteilen', 
    description: 'Upload folder containing yet to be organized files',
    input_type: 'text'
  },
  'books_folder': {
    value: 'Books', 
    description: 'Root folder of Library.\nEnter <b>unique name</b> or <b>folder id</b>\n(will be changed into ID by the addon).',
    input_type: 'text'
  }
}


/**
 * 
 * Event handler
 * 
 */

/**
 * Callback for rendering the homepage card.
 * @return {CardService.Card} The card to show to the user.
 */
function onHomepage(e) {

  return createInstructionsCard();
}

/**
 * Callback for rendering card when an item is selected. Depending on the selection
 *    - instructions card
 *    - book organizer card
 *    - book display card
 * are returned.
 * 
 * @return {CardService.Card} The card to show to the user.
 */
function onDriveItemsSelected(e) {
  var item_id = e.drive.activeCursorItem.id,
    item_title = e.drive.activeCursorItem.title,
    mimeType = e.drive.activeCursorItem.mimeType;

  var fregex = new RegExp(getSetting('special_folder_name'), 'gim');
  if (mimeType.match(/folder/gim) && item_title.match(fregex)){
    //we are in the special folder, show organizer card
    return createBookOrganizerCard(item_id, item_title);
  } else if (mimeType.match(/pdf|epub/gim)){
    //epub or pdf selected, return display book card
    return createDisplayBookCard(item_id, item_title);
  } else {
    //neither selected, return instructions card
    return createInstructionsCard();
  }
}

/**
 * Callback for rendering card when a user has selected a book in the book organizer card.
 * 
 * @return {CardService.Card} The card to show to the user.
 */
function onDisplayBook(e) {
  //build response card
  var card = createDisplayBookCard(e.commonEventObject.parameters.book_id, e.commonEventObject.parameters.book_name)

  // Create an action response that instructs the add-on to replace
  // the current card with the new one.
  var navigation = CardService.newNavigation()
      .pushCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();

}

/**
 * Callback for editing a book from the book display card.
 * Re-opens the display card afterwards
 * 
 * @return {CardService.Card} The card to show to the user.
 */
function onEditBook(e){
  var book_name = e.commonEventObject.parameters.book_name,
    move = e.commonEventObject.parameters.move,
    book_id = e.commonEventObject.parameters.book_id,
    title = (e.commonEventObject.formInputs.title ? e.commonEventObject.formInputs.title.stringInputs.value[0] : false),
    author = (e.commonEventObject.formInputs.author ? e.commonEventObject.formInputs.author.stringInputs.value[0] : false),
    year = (e.commonEventObject.formInputs.year ? e.commonEventObject.formInputs.year.stringInputs.value[0] : false);

  // show error notification if either title or author are empty
  if (!title || !author) {
    var flds = ['The field(s)'];
    if (!title)
      flds.push('title');
    if (!author)
      flds.push('author');
    var text = "The field(s) ${flds.join(',')} cannot be empty";

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(text))
      .build();
  } 

  author = author.trim();
  title = title.trim();
 
  //STEP 1: rename file
  var file = DriveApp.getFileById(book_id);
  var file_ending = file.getMimeType().match(/pdf/gi) ? '.pdf' : '.epub';
  var fn_parts = [title, "["+author+"]"];
  if (year) {
     year = year.trim();
    fn_parts.push("("+year+")");
  }
  var new_filename = fn_parts.join('_');
  new_filename += file_ending;
  file.setName(new_filename);

  //STEP 2: move file if necessary
  if(move && move == 'on'){
    var books_folder = getFolderFromSetting('books_folder');
    var destination = books_folder.getFoldersByName(author);
    if (!destination.hasNext()){
      destination = DriveApp.createFolder(author);
      destination.moveTo(books_folder);
    } else {
      destination = destination.next();
    }
    if (file.getParents().next() != destination){
      file.moveTo(destination);
    }
  }

  var card = createDisplayBookCard(book_id, new_filename);

  // Create an action response that instructs the add-on to replace
  // the current card with the new one.
  var navigation = CardService.newNavigation()
      .updateCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();  
}

function onGetMetadata(e){
  var card = createDisplayBookCard(e.commonEventObject.parameters.book_id, e.commonEventObject.parameters.book_name, true);

  var navigation = CardService.newNavigation()
      .pushCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}

/**
 * Create a collection of cards to control the add-on settings and
 * present other information. These cards are displayed in a list when
 * the user selects the associated "Open settings" universal action.
 *
 * @param {Object} e an event object
 * @return {UniversalActionResponse}
 */
function onOpenSettings(e) {
  return CardService.newUniversalActionResponseBuilder()
      .displayAddOnCards(
          [createSettingsCard(), createAboutCard()])
      .build();
}

function onSaveSettings(e) {
  var card = createSettingsCard('Settings have been updated.');

  var navigation = CardService.newNavigation()
      .pushCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}

function onResetSettings(e){
  let settings = {};
  for(const key in default_settings) {
    settings[key] = default_settings[key].value;
  }
  //console.log(settings);
  PropertiesService.getUserProperties().setProperties(settings);

  var card = createSettingsCard('Settings have been reset.');
  var navigation = CardService.newNavigation()
      .updateCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}
/**
 * Run background tasks, none of which should alter the UI.
 * Also records the time of sync in the script properties.
 *
 * @param {Object} e an event object
 */
function onCreateStatistics(e) {
  createStatistics();
  // no return value tells the UI to keep showing the current card.
  return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Statistics have been updated. Details can be found in the Library Sheet'))
      .build();
}

/**
 * Callback for rendering card when a user has selected a book in the book organizer card.
 * 
 * @return {CardService.Card} The card to show to the user.
 */
function onOpenStatistics(e) {
  //build response card
  var card = createStatisticsCard();

  // Create an action response that instructs the add-on to replace
  // the current card with the new one.
  var navigation = CardService.newNavigation()
      .pushCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();

}

function onListNonConformantBooks(){
  var card = createNonConformantBooksCard();

  var navigation = CardService.newNavigation()
      .pushCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build(); 
}


/**
 * 
 * Card factories
 * 
 */


/**
 * Create and return a built settings card.
 * @return {Card}
 */
function createSettingsCard(msg) {
  let props = PropertiesService.getUserProperties();
  let missing_settings = {};

  //make sure all settings have been created, if not add them now
  for(const setting in default_settings) {
    if (!props.getKeys().includes(setting)){
      missing_settings[setting] = default_settings[setting].value;
    }
  }
  if (missing_settings.keys){
    props.setProperties(missing_settings);
    props = PropertiesService.getUserProperties();
  }

  // CREATE GUI
  let cb = CardService.newCardBuilder();
  // Add message if available
  if(msg){
    cb.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText('<b><font color="#0e2cc0">'+msg+'</font></b>') 
    ))
  }
  // Show settings
  for(k in default_settings){
    let title = k.replace(/_/g,' ').replace(/(^|\s)\S/g, function(t) { return t.toUpperCase() });
    switch (default_settings[k].input_type) {
    case 'text':
      cb.addSection(CardService.newCardSection()
        .addWidget(CardService.newTextInput().setTitle(title).setFieldName(k).setValue(props.getProperty(k)))
        .addWidget(CardService.newTextParagraph().setText(default_settings[k].description))
      )
      break;
    case 'textarea':
      cb.addSection(CardService.newCardSection()
        .addWidget(CardService.newTextInput().setMultiline(true).setFieldName(k).setTitle(title).setValue(props.getProperty(k)))
        .addWidget(CardService.newTextParagraph().setText(default_settings[k].description))
      )
      break;  
    case 'none':
      cb.addSection(CardService.newCardSection()
        .addWidget(CardService.newTextInput().setTitle(title).setFieldName(k).setValue(props.getProperty(k)))
        .addWidget(CardService.newTextParagraph().setText(default_settings[k].description))
      )
      break;
    }
  }

  //section = CardService.newCardSection().addWidget(CardService.newTextParagraph().setText('test'));

  return cb.setName('settings').setHeader(CardService.newCardHeader().setTitle('Settings'))
      .setFixedFooter(
        CardService.newFixedFooter().setPrimaryButton(
          CardService.newTextButton().setText('Save').setOnClickAction(
            CardService.newAction().setFunctionName('onSaveSettings')
          )
      ).setSecondaryButton(
          CardService.newTextButton().setText('Reset').setOnClickAction(
            CardService.newAction().setFunctionName('onResetSettings')
          )        
      )).build();
}

/**
 * Create and return a built 'About' informational card.
 * @return {Card}
 */
function createAboutCard() {
  var cs = CardService.newCardBuilder()
    .setName('about')
    .setHeader(CardService.newCardHeader().setTitle('About'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(`This add-on manages an eBook library. It currently has support for PDF and ePub books.
        Features include facilities to rename books based on metadata, building statistics on the library.`)
      )
    )

  return cs.build();  // Don't forget to build the card!
}

function createStatisticsCard() {
  let props = PropertiesService.getUserProperties();

  var cs = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Statistics'))
      .setName('statistics');

  if (props.getProperty('num_books')){
    cs.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('The library contains <b>'+props.getProperty('num_folders')+'</b> Authors / Folders.')
      )
    );
    cs.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('The library contains <b>'+props.getProperty('num_books')+'</b> Books.')
      )
    );    
  }
  if (props.getProperty('stat_sheet')){
    var button = CardService.newTextButton()
      .setText("Stats Sheet")
      .setOpenLink(CardService.newOpenLink()
        .setUrl("https://docs.google.com/spreadsheets/d/"+props.getProperty('stat_sheet'))
        .setOpenAs(CardService.OpenAs.FULL_SIZE)
        .setOnClose(CardService.OnClose.NOTHING));
    cs.addSection(CardService.newCardSection()
      .addWidget(CardService.newDecoratedText()
        .setText('More information can be found in this file')
        .setButton(button)
      )
    );
  } else {
    cs.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(`More throrough stats are available once "Create Statistics" has been selected from the Addon-Menu.
        Note that this will index 50 books per call and will have to be called more than once for large libraries.`)
      )
    );
  }

  cs.setFixedFooter(CardService.newFixedFooter().setPrimaryButton(
      CardService.newTextButton().setTextButtonStyle(CardService.TextButtonStyle.OUTLINED).setText('Update Statistics').setOnClickAction(
        CardService.newAction().setFunctionName('createStatistics')
      )
    )    
  )

  return cs.build();  // Don't forget to build the card!
}



/**
 * Create the instructions card shown on the homepage or when nothing relevant is selected.
 * 
 * @return {CardService.Card} The card to show to the user.
 */
function createInstructionsCard(){
  var text_section = CardService.newCardSection();
  text_section.addWidget(CardService.newTextParagraph().setText('Select a book file or the folder containing books to organize.'));
  text_section.addWidget(CardService.newTextParagraph().setText('The folder for books to organize must be named: "'+getSetting('special_folder_name')+'".'));

  button_section = CardService.newCardSection()
  .addWidget(
    CardService.newTextButton()
      .setText('List non-conformant books')
      .setOnClickAction(CardService.newAction().setFunctionName('onListNonConformantBooks')
    )
  )
  var card = CardService.newCardBuilder()
      .setName('home')
      .addSection(text_section)
      .addSection(button_section);

  return card.build();
}

/**
 * Creates the book organizer card listing files in a special folder.
 * @param folder_id {String} The id of the special folder.
 * 
 * @return {CardService.Card} The assembled card.
 */
function createBookOrganizerCard(folder_id) {
  var section = CardService.newCardSection(),
    texts = [],
    files = [],
    peekHeader;
  var folder = DriveApp.getFolderById(folder_id); 
  var files = folder.getFiles();

  if (!files.hasNext()){
    // folder is empty: Text: Folder is empty, nothing to organize
    section.addWidget(CardService.newTextParagraph().setText('Folder is empty, nothing to organize.'));
  } else {
    // folder contains files: List of files and button: edit
    section.addWidget(CardService.newTextParagraph().setText('Select book and click "Edit".'));
  
    while (files.hasNext()){
      file = files.next();

      let act = CardService.newAction().setFunctionName('onDisplayBook').setParameters({"book_name": file.getName(), "book_id": file.getId()});
      let btn = CardService.newTextButton().setText('Edit').setOnClickAction(act);
      let dt = CardService.newDecoratedText()
        .setText(file.getName())
        .setBottomLabel('Current title')
        .setWrapText(true)
        .setButton(btn);
      
      section.addWidget(dt)
    }
  }

  var card = CardService.newCardBuilder()
      .setName('organizer')
      .addSection(section);

  return card.build();
}

function createNonConformantBooksCard(){
  let section = CardService.newCardSection(),
    folders,
    books,
    books_to_list = [];

  [folders, books] = getLibraryData();

  for(book of books){
    if (!book.name.match(standard_filter)){
      books_to_list.push(book);
    }
    if(books_to_list.length > 10){
      break;
    }
  }
  if (books_to_list.length == 0){
    // no books to edit, nothing to organize
    section.addWidget(CardService.newTextParagraph().setText('All books are conformant, nothing to do.'));
  } else {
    // folder contains files: List of files and button: edit
    section.addWidget(CardService.newTextParagraph().setText('Select book and click "Edit".'));
  
    for(book of books_to_list){
      let act = CardService.newAction().setFunctionName('onDisplayBook').setParameters({"book_name": book.name, "book_id": book.id});
      let btn = CardService.newTextButton().setText('Edit').setOnClickAction(act);
      let dt = CardService.newDecoratedText()
        .setText(book.name)
        .setBottomLabel('Current title')
        .setWrapText(true)
        .setButton(btn);
      
      section.addWidget(dt);

      let parent = DriveApp.getFileById(book.id).getParents().next();
      section.addWidget(CardService.newDecoratedText()
        .setText('In Folder: '+parent.getName())
        .setBottomLabel('Opens in new tab')
        .setWrapText(true)
        .setButton(CardService.newTextButton()
          .setText('Open').setOpenLink(CardService.newOpenLink().setUrl(parent.getUrl())
        )
      )); 
      section.addWidget(CardService.newDivider());     
    }
  }  
  let card = CardService.newCardBuilder()
      .setName('non-conformant')
      .addSection(section);

  return card.build();

}

/**
 * Creates the book edit card show form fields for the book title.
 * @param book_id {String} The id of the book.
 * @param book_name {String} The current name of the book.
 * @param incl_meta {boolean} Whether or not the metadata of the book should be shown
 * 
 * @return {CardService.Card} The assembled card.
 */
function createDisplayBookCard(book_id, book_name, incl_meta){
  var bsfld = getFolderFromSetting('books_folder');
  var aflds = bsfld.getFolders();
  var authors = [];

  while (aflds.hasNext()) {
    authors.push(aflds.next().getName().trim());
  };

  //rubric section
  var rubric_section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
    .setText(`Set author, title and (optionally) year book was published.
    When selecting "Save and Move" the book will be renamed and moved into the respective author folder.\nIf the author folder does not exist, it will be created.
    When selecting "Save" the book will be renamed but remain in the folder.
    Select "Show Metadata to scrape file metadata of the selected book.`)
  )

  //form section
  var form_section = CardService.newCardSection();
  var {title, author, year} = extractMetadataFromFileName(book_name);

  let t_input =CardService.newTextInput()
    .setFieldName('title')
    .setHint('Title from filename')
    .setTitle('Enter new title');
  if (title){
    t_input.setValue(title);
  }
  form_section.addWidget(t_input);
  form_section.addWidget(CardService.newDivider());

  let validation_text = CardService.newTextParagraph();
  let a_input =CardService.newTextInput()
    .setFieldName('author')
    .setTitle('Enter new author')
    .setSuggestions(CardService.newSuggestions().addSuggestions(authors));
  if (author && authors.includes(author)){
    validation_text.setText('<font color="#1fa432">Author exists, Folder exists</font>')
    a_input.setValue(author);
  } else if (author){
    validation_text.setText('<font color="#ff7200">Author from filename</font>')
    a_input.setValue(author);
  } else {
    validation_text.setText('<font color="#ff0000">Author and Title are required</font>')
  }
  form_section.addWidget(a_input);
  form_section.addWidget(validation_text);
  form_section.addWidget(CardService.newDivider());

  let y_input =CardService.newTextInput()
    .setFieldName('year')
    .setTitle('Enter new yaer');
  if (year){
    y_input.setValue(year).setHint('Year from filename');
  }
  form_section.addWidget(y_input);
  form_section.addWidget(CardService.newDivider());

  //metadata section
  var metadata_section = CardService.newCardSection();
  if (incl_meta){
    var book = DriveApp.getFileById(book_id);
    var metadata = book.getMimeType().match(/pdf/gim) ? getPdfMetadata(book) : getEpubMetadata(book);

    if (Object.keys(metadata).length > 0){
      for(key in metadata){
        metadata_section.addWidget(CardService.newTextParagraph()
          .setText(`<b>${key}</b>\n`+metadata[key])
        ).addWidget(CardService.newDivider());
      }
    } else {
      metadata_section.addWidget(CardService.newTextParagraph()
          .setText('<b>No metadata found.</p>')
        ).addWidget(CardService.newDivider());      
    }
  } else {
    metadata_section.addWidget(CardService.newDivider());
    metadata_section.addWidget(CardService.newDecoratedText()
      .setText('Reloads Page')
      .setButton(CardService.newTextButton()
        .setText('Show Metadata')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('onGetMetadata')
          .setParameters({"book_name": book_name, "book_id": book_id})
        )
      )
    )
  }
 
  //footer
  let editAction = CardService.newAction().setFunctionName('onEditBook');
  let moveButton = CardService
    .newTextButton()
    .setText("Save and Move")
    .setOnClickAction(editAction.setParameters({"book_name": book_name, "book_id": book_id, "move": "on"}));
  let saveButton = CardService
    .newTextButton()
    .setText("Save")
    .setOnClickAction(editAction.setParameters({"book_name": book_name, "book_id": book_id, "move": "off"}));


  var footer = CardService.newFixedFooter()
    .setPrimaryButton(saveButton)
    .setSecondaryButton(moveButton);

  var cs = CardService.newCardBuilder()
    .setName('display')
    .setHeader(CardService.newCardHeader().setTitle("Edit Book"))
    .addSection(rubric_section)
    .addSection(form_section)
    .addSection(metadata_section)
    .setFixedFooter(footer);

  //assemble card
  return cs.build();
}

/**
 * 
 * Helper functions
 * 
 */

function chunk(arr, chunkSize) {
  if (chunkSize <= 0) throw "Invalid chunk size";
  var R = [];
  for (var i=0,len=arr.length; i<len; i+=chunkSize)
    R.push(arr.slice(i,i+chunkSize));
  return R;
}

function getSetting(setting){
  var props = PropertiesService.getUserProperties();
  var ps = props.getProperty(setting);
  if (!ps){
    ps = default_settings[setting].value;
    props.setProperty(setting, ps);
  } 
  return ps;
}

function getFolderFromSetting(setting){
  var setting = getSetting(setting);
  var folder;
  if (setting.match(/^\d[\w_-]{25,}$/gmi)){
    folder = DriveApp.getFolderById(setting);
  } else {
    var folder = DriveApp.getFoldersByName(setting).next();
    PropertiesService.getUserProperties().setProperty(setting, folder.getId());
  } 
  return folder;
}

/**
 * Retrieves the metadata from the PDF.
 * Uses PDF-lib (https://pdf-lib.js.org)
 * @param book {File} The book.
 * 
 * @return {Object} The retrieved metadata.
 */
function getPdfMetadata(book) {
  var blob = book.getBlob();

  var scraper = new PDFScraper(blob.getDataAsString());
  var metadata = scraper.pretty_print_unique_values('JS');

  // for(md in metadata){
  //   metadata[md] = metadata[md].replace(/\p{C}|\p{P}|\p{S}/gu, '').replace(/\n\r|\p{Z}/gu, ' ').replace(/\/\\/g, '');
  // }

  return metadata;
}

/**
 * Retrieves the metadata from the ePub.
 * Inspired by epub-metadata-parser (https://github.com/aeroith/epub-metadata-parser)
 * @param book {File} The book.
 * 
 * @return {Object} The retrieved metadata.
 */
function getEpubMetadata(book) {
  var epub = book.getBlob();
  epub.setContentType(MimeType.ZIP);
  var files = Utilities.unzip(epub);

  var opf_file = files.find(function(f){
    if(f.getName().match(/opf$/gim))
      return true;
    else
      return false;
  })

  let opf = XmlService.parse(opf_file.getDataAsString());
  let r_elem = opf.getRootElement();
  var ns = XmlService.getNamespace("http://www.idpf.org/2007/opf");

  let m_elem = r_elem.getChild('metadata', ns);
  let metadata = {}; 
  m_elem.getChildren().forEach( item => {
    metadata[item.getName()] = item.getText();
  });

  return metadata;
}

/**
 * Retrieves the metadata from the file name of the book.
 * Iterates over different filters defined in METADATA_FILTER
 * @param book {String} The book name.
 * 
 * @return {Array} The retrieved metadata.
 */
function extractMetadataFromFileName(book_name){
  var metadata = {'title': false, 'author': false, 'year': false},
    match;
  
  //remove file ending
  book_name = book_name
    .replace(/\.[\w]{3,4}/ig, '')
    .trim();

  //apply filters to find title, author, year, from first to last and breaks the first time a match is found
  for (filter of metadata_filter){
    if(match = book_name.match(filter.regex)){
      //console.log('regex: '+filter.regex.toString()+' match: '+JSON.stringify(match))
      filter.fields.forEach(function(field, pos){
        //to compensate for full match add +1 to get to the capture groups
        metadata[field] = match[pos + 1] ? match[pos + 1].replace(/[\p{P}\s]+/giu, ' ').trim() : false;
      })
      break;
    }
  };

  //ensure, title is always set to something useful
  if(!metadata.title){
    metadata.title = book_name.replace(/[\p{P}\s]+/giu, ' ').trim();
  } 

  return metadata;
}

function getLibraryData(){
  var books_folder = getFolderFromSetting('books_folder');
  var folders = Drive.Files.list({q: `mimeType = 'application/vnd.google-apps.folder' and '${books_folder.getId()}' in parents  and trashed = false`, pageSize: 500});
  var book_search;
  var books_chunk;
  var books = [];

  var chunks = chunk(folders.files, 100);
  
  for(let i=0; i < chunks.length; i++){
    book_search = [];
    chunks[i].map(function(f){
      book_search.push(`'${f.id}' in parents`);
    });
    books_chunk = Drive.Files.list({q:`mimeType != 'application/vnd.google-apps.folder' and trashed = false and ( ${book_search.join(' or ')} )`, pageSize: 500});
    books = books.concat(books_chunk.files);
  }
  
  return [folders.files, books];
}

function createStatistics(stat_sheet){
  var books = [],
      books_sheet,
      folder_sheet,
      props,
      folders = [],
      stat_sheet_id,
      stat_sheet;

  //initialize or get stat_sheet
  props = PropertiesService.getUserProperties();
  if (!(stat_sheet_id = props.getProperty('stat_sheet'))){
    stat_sheet = SpreadsheetApp.create("Library Statistics [generated]");
    props.setProperty('stat_sheet', stat_sheet.getId());
  } else {
    stat_sheet = SpreadsheetApp.openById(stat_sheet_id);
  }      

  //get library data and update stats in properties
  [folders, books] = getLibraryData();
  props = PropertiesService.getUserProperties();
  props.setProperties({'num_folders': folders.length, 'num_books': books.length});

  //prepare google sheet
  if (!stat_sheet.getSheetByName('folders')){
    folder_sheet = stat_sheet.insertSheet('folders');
  } else {
    folder_sheet = stat_sheet.getSheetByName('folders');
  };
  if (!stat_sheet.getSheetByName('books')){
    books_sheet = stat_sheet.insertSheet('books');
  } else {
    books_sheet = stat_sheet.getSheetByName('books');
  };  

  //reset folders sheet and write data to sheet
  folder_sheet.clear();
  folder_sheet.getRange(1,1,1,2).setValues([['Folder ID',	'Folder Name']]).setFontWeight("bold");
  folder_sheet.setFrozenRows(1);
  folder_sheet.getRange(2, 1, folders.length, 2).activate().setValues(folders.map(f => ([f.id, f.name])));

  //reset books sheet and write data to sheet
  books_sheet.getRange(2,1,books_sheet.getLastRow(),6).clear();
  books_sheet.clear();
  books_sheet.getRange(1,1,1,5).setValues([['Folder ID',	'Folder Name',	'Book ID',	'Book Name',	'Conformant?']]).setFontWeight("bold");
  books_sheet.setFrozenRows(1);  
  books.forEach(function(b,idx){
    books[idx].up_to_date = b.name.match(standard_filter) ? 'yes' : 'no'
  })
  books_sheet.getRange(2, 3, books.length, 3).activate().setValues(books.map(f => ([f.id, f.name, f.up_to_date])));
  
}

var timer = {
    start: Date.now(),
    since: function(){
        return 'time since: '+Math.round((Date.now() - this.start)/1000);
    },
    reset: function(){
        this.start = Date.now();
    }
}
