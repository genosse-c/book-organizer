
/**
 * Runs a background task creating statistics that are stored in user properties and
 * in a google sheet. Shows a notification to the user upon completion
 *
 * @param {Object} e an event object
 * @return {CardService.ActionResponse} A notification to show to the user.
 */
function onCreateStatistics(e) {
  createStatistics();

  return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification()
          .setText('Statistics have been updated. Details can be found in the Library Sheet')
        )
      .build();
}

/**
 * Callback for rendering the statistics card.
 *
 * @param {Object} e an event object
 * @return {CardService.ActionResponse} The card to show to the user.
 */
function onOpenStatistics(e) {
  //build response card
  const card = createStatisticsCard();

  let navigation = CardService.newNavigation()
      .pushCard(card);
  let actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();

}

/**
 * Creates the statistics card.
 *
 * @return {CardService.Card} The card to show to the user.
 */
function createStatisticsCard() {
  const props = PropertiesService.getUserProperties();

  let cs = CardService.newCardBuilder()
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
    let button = CardService.newTextButton()
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
        Note that this might have to be called more than once for large libraries.`)
      )
    );
  }

  cs.setFixedFooter(CardService.newFixedFooter().setPrimaryButton(
      CardService.newTextButton().setTextButtonStyle(CardService.TextButtonStyle.OUTLINED).setText('Update Statistics').setOnClickAction(
        CardService.newAction().setFunctionName('onCreateStatistics')
      )
    )
  )

  return cs.build();  // Don't forget to build the card!
}

/**
 * Helper function: Identifies all files and folders in the library and returns their descriptors
 *
 * @return {array} An array holding an array of folder descriptors and an array of file descriptors
 */
function getLibraryData(){
  const books_folder = getFolderFromSetting('books_folder');
  const folders = Drive.Files.list({q: `mimeType = 'application/vnd.google-apps.folder' and '${books_folder.getId()}' in parents  and trashed = false`, pageSize: 500});
  let book_search,
    books_chunk,
    books = [];

  // split the folders into chunks of 100 in order to not breach the size limit of the query string
  // that is used to identify the book files
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

/**
 * Finds or initializes the spreadsheet containing statistics on the eBook library
 * and writes folder and book file information to it.
 * Information in the Spreadsheet is overwritten each time this function is called
 * In order to be useful, this information has to be postprocessed in the library spreadsheet
 */
function createStatistics(){
  let books_sheet,
      folder_sheet,
      stat_sheet_id,
      stat_sheet;

  const props = PropertiesService.getUserProperties();
  //initialize or get stat_sheet
  if (!(stat_sheet_id = props.getProperty('stat_sheet'))){
    stat_sheet = SpreadsheetApp.create("Library Statistics [generated]");
    props.setProperty('stat_sheet', stat_sheet.getId());
  } else {
    stat_sheet = SpreadsheetApp.openById(stat_sheet_id);
  }

  //get library data and update stats in properties
  const [folders, books] = getLibraryData();
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

