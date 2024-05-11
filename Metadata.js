// array of objects each with
// a *regex* property to find up to three parts of name metadata in the filename
// a *fields* property consisting of an array naming the fields: "title", "author", "year"
// the filters should sorted from highest specificity to lowest
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

/**
 * Callback for retrieving statistics to be displayed in the Display Book card.
 *
 * @param {Object} e an event object
 * @return {CardService.ActionResponse} The card to show to the user.
 */
function onGetMetadata(e){
  const card = createDisplayBookCard(e.commonEventObject.parameters.book_id, e.commonEventObject.parameters.book_name, true);

  let navigation = CardService.newNavigation()
      .pushCard(card);
  let actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}

/**
 * Retrieves the metadata from the PDF.
 * Uses pdf-metadata-scraper (https://github.com/genosse-c/pdf_metadata_scraper)
 *
 * @param book {DriveApp.File} The book to be scraped
 * @return {Object} The retrieved metadata.
 */
function getPdfMetadata(book) {
  const blob = book.getBlob();

  const scraper = new PDFScraper(blob.getDataAsString());
  let metadata = scraper.pretty_print_unique_values('JS');

  return metadata;
}

/**
 * Retrieves the metadata from the ePub.
 * Inspired by epub-metadata-parser (https://github.com/aeroith/epub-metadata-parser)
 *
 * @param book {DriveApp.File} The book.
 * @return {Object} The retrieved metadata.
 */
function getEpubMetadata(book) {
  const epub = book.getBlob();
  epub.setContentType(MimeType.ZIP);
  let files = Utilities.unzip(epub);

  const opf_file = files.find(function(f){
    if(f.getName().match(/opf$/gim))
      return true;
    else
      return false;
  })

  const opf = XmlService.parse(opf_file.getDataAsString());
  const r_elem = opf.getRootElement();
  const ns = XmlService.getNamespace("http://www.idpf.org/2007/opf");

  const m_elem = r_elem.getChild('metadata', ns);
  let metadata = {};
  m_elem.getChildren().forEach( item => {
    metadata[item.getName()] = item.getText();
  });

  return metadata;
}

/**
 * Retrieves the metadata from the file name of the book.
 * Iterates over different filters defined in METADATA_FILTER
 * @param book_name {String} The book name.
 *
 * @return {Array} The retrieved metadata.
 */
function extractMetadataFromFileName(book_name){
  let metadata = {'title': false, 'author': false, 'year': false},
    match;

  //remove file ending
  book_name = book_name
    .replace(/\.[\w]{3,4}/ig, '')
    .trim();

  //apply filters to find title, author, year, from first filter to last and breaks the first time a match is found
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

