//Tests whether the book name follows the "{title}_[{author}]_({year})" format with year being optional
const standard_filter = /^([^\[]+)[_]*\[([^\]]+)]/m 

/**
 * Callback for rendering the homepage card.
 * @param {Object} e an event object
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
 * @param {Object} e an event object
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
 * Callback for rendering a list of 10 non-conformant books
 * the card can be accessed from the instructions card.
 *
 * @param {Object} e an event object
 * @return {CardService.Card} The card to show to the user.
 */
function onListNonConformantBooks(){
  const card = createNonConformantBooksCard();

  let navigation = CardService.newNavigation()
      .pushCard(card);
  let actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build(); 
}

/**
 * Create the instructions card shown on the homepage or when nothing relevant is selected.
 * 
 * @return {CardService.Card} The card to show to the user.
 */
function createInstructionsCard(){
  let text_section = CardService.newCardSection();
  text_section.addWidget(CardService.newTextParagraph().setText('Select a book file or the folder containing books to organize.'));
  text_section.addWidget(CardService.newTextParagraph().setText('The folder for books to organize must be named: "'+getSetting('special_folder_name')+'".'));

  let button_section = CardService.newCardSection()
  .addWidget(
    CardService.newTextButton()
      .setText('List non-conformant books')
      .setOnClickAction(CardService.newAction().setFunctionName('onListNonConformantBooks')
    )
  )
  let card = CardService.newCardBuilder()
      .setName('home')
      .addSection(text_section)
      .addSection(button_section);

  return card.build();
}

/**
 * Creates the book organizer card listing files in a special folder.
 * It is shown when the special folder is selected in the drive
 *
 * @param folder_id {String} The id of the special folder.
 * 
 * @return {CardService.Card} The assembled card.
 */
function createBookOrganizerCard(folder_id) {
  let section = CardService.newCardSection(),
    texts = [],
    files = [],
    peekHeader;
  const folder = DriveApp.getFolderById(folder_id);
  let files = folder.getFiles();

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

  let card = CardService.newCardBuilder()
      .setName('organizer')
      .addSection(section);

  return card.build();
}

/**
 * Creates a card listing 10 that do not conform to the naming specifications
 *
 * @return {CardService.Card} The assembled card.
 */
function createNonConformantBooksCard(){
  let section = CardService.newCardSection(),
    books_to_list = [];

  const [folders, books] = getLibraryData();

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
    // non conformant files found: List the files and add a button to edit or go to the drive folder
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
 * Helper function: Splits an array into multiple arrays containing a specified number of elements of the source array.
 *
 * @param arr {array} The array to be split into chunks
 * @param chunksize {Integer} The size of each chunk
 * 
 * @return {array} An array of arrays containing the source array in chunks
 */
function chunk(arr, chunkSize) {
  if (chunkSize <= 0) throw "Invalid chunk size";
  let chunks = [];
  for (let i=0; i<arr.length; i+=chunkSize)
    chunks.push(arr.slice(i,i+chunkSize));
  return chunks;
}

