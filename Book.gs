
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
