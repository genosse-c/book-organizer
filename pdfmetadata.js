/*
 * PDFScraper
 *
 * Copyright 2024 Conrad Noack
 *
 * May 4th 2024
 *
 * This software relies heavily on https://github.com/preciz/pdf_info (MIT License)
 * with some additional inspiration based on https://github.com/sindresorhus/uint8array-extras (MIT License)
 */


class PDFScraper {
  /**
  * Constructor for PDFScraper
  *
  * @param  binary variable holding the binary representation of the PDF to be scrapped
  * @return      instantiated JSGloss object
  */
  constructor(binary) {
    this.binary = binary;
  }

  /**********************
  * main public functions
  ***********************/

  /**
  * Outputs the found metadata as json
  *
  * @type {string}   if either meta or info, will only scrape metadata or info objects respectivly
  * @return {string} pretty printed json represenation of metadata items from the PDF
  */
  pretty_print(type){
    let data;
    switch(type){
      case "meta":
        data = this.metadata_objects();
        break;
      case "info":
        data = this.info_objects();
      default:
        data = {metadata: this.metadata_objects(), info: this.info_objects()};
    }
    return JSON.stringify(data, null, '  ');
  }

  /**
  * Outputs found metadata but reduced to those with unique values either as JS, JSON or as HTML
  *
  * @markup   switch to select the type of output. Currently supports JS, JSON string, HTML
  * @return {string} json or HTML represenation of metadata items from the PDF
  */
  pretty_print_unique_values(format){
    let info = this.info_objects();
    let meta = this.metadata_objects();
    let data = [info, meta];

    this.unique = {};
    this.walk_data(data);
    console.log(this.unique);
    switch (format) {
      case 'JSON':
        return JSON.stringify(this.unique, null, '  ');
      case 'JS':
        return {...this.unique}
      case 'HTML':
        var keyArray = Object.keys(this.unique);
        keyArray.sort((a, b) => a.length - b.length);
        let html = '<dl class="pdf-scraper">';
        keyArray.forEach(function(key){
          html += `<dt>${key}</dt><dd>${this.unique[key]}</dd>`
        }, this)
        html += '</dl>';
        return html;
    }

  }
  /**
  * Helper function to reduce found metadata to unique metadata values
  *
  * @data   data to be reduced
  */
  walk_data(data){
    if (Array.isArray(data)){
      data.forEach(function(elem){
        this.walk_data(elem);
      }, this);
    } else if (typeof data === 'object' && data !== null){
      Object.keys(data).forEach(function(key){
        if (typeof data[key] === 'string'){
          if (!Object.values(this.unique).includes(data[key])){
             this.unique[key] = data[key];
          }
        } else {
          this.walk_data(data[key]);
        }
      }, this);
    } else {
       console.log('we should not be here: '+JSON.stringify(data));
    }
  }

  /**
  * Returns an object containing all metadata items stored in info objects in the PDF
  * Metadata values are decoded
  * @return {object} Object holding the metadata
  */
  info_objects() {
    const refs = this.info_refs();
    const objects = {};

    if (refs){
      for (const ref of refs) {
        const objId = ref.slice(6, -2);
        const obj = this.get_object(objId);
        if (obj){
          const list = obj.flat().filter((item, index, self) => self.indexOf(item) === index);
          objects[ref] = list.map(rawInfoObj => this.parse_info_object(rawInfoObj));
        }
      }
    }

    if (Object.keys(objects).length === 0){
      return {'error': 'no info objects found'}
    } else {
      return objects
    }
  }

  /**
  * Returns an object containing all metadata items stored in metadata objects in the PDF
  * Metadata values are decoded
  * @return {object} Object holding the metadata
  */
  metadata_objects() {
    const rawMetadataObjects = this.raw_metadata_objects()
    const objects = [];
    for (const rawMetadata of rawMetadataObjects) {
      const map = this.parse_metadata_object(rawMetadata);
      if (map && Object.keys(map).length > 0) {
        objects.push(map);
      }
    }
    return objects;
  }

  /**********************
  * raw output functions
  ***********************/

  /**
  * Returns an object containing all metadata items stored in info objects in the PDF
  * Metadata values are NOT decoded
  * @return {object} Object holding the raw metadata
  */
  raw_info_objects() {
    const refs = this.info_refs();
    const objects = {};
    for (const ref of refs) {
      const objId = ref.slice(6, -2);
      const obj = this.get_object(objId);
      const list = obj.flat().filter((item, index, self) => self.indexOf(item) === index);
      objects[ref] = list;
    }
    return objects;
  }

  /**
  * Returns an object containing all metadata items stored in metadata objects in the PDF
  * Metadata values are NOT decoded
  * @return {object} Object holding the raw metadata
  */
  raw_metadata_objects() {
    const regexStart = /<x:xmpmeta/g;
    const regexEnd = /<\/x:xmpmeta/g;
    const starts = [...this.binary.matchAll(regexStart)].map(match => match.index);
    const ends = [...this.binary.matchAll(regexEnd)].map(match => match.index);
    const pairs = [];
    for (const start of starts) {
      for (const end of ends) {
        if (start < end) {
          pairs.push([start, end]);
          break;
        }
      }
    }
    const rawMetadataObjects = pairs.map(pair => this.binary.slice(pair[0], pair[1] + 12));
    return rawMetadataObjects;
  }

  /**********************
  * analytic functions
  ***********************/

  /**
  * Determines whether the variable handed to the constructor is actually a PDF
  * @return {Boolean} is PDF?
  */
  is_pdf(){
    if (this.binary.startsWith("%PDF-")) {
      return true;
    }

    const head = this.binary.slice(0, Math.min(1024, this.binary.length));
    return head.includes("%PDF-");
  }

  /**
  * Determines the PDF version
  * @return {Object} PDF version
  */
  pdf_version(bin) {
    const data = bin ? bin : this.binary;

    if (data.startsWith("%PDF-")) {
      const version = data.slice(5, 8);
      return { ok: version };
    }

    const head = data.slice(0, Math.min(1024, data.length));
    const match = head.match(/%PDF-[0-9]\.[0-9]/);
    if (match) {
      return this.pdf_version(match[0]);
    } else {
      return { error: true };
    }
  }
  /**
  * Determines wether the PDF version
  * @return {Boolean} Is PDF encrypted?
  */
  is_encrypted() {
    const refs = this.encrypt_refs();
    return refs.length > 0;
  }

  /***************************************
  * private functions, that
  * should not be called from the outside
  ***************************************/

  /**********************
  * metadata wrangling
  ***********************/

  parse_metadata_object(string) {
    if (typeof string === 'string') {
      const xmpMatch = string.match(/<x:xmpmeta.*<\/x:xmpmeta>/gsm);
      if (xmpMatch) {
        const xmp = xmpMatch[0];
        const tags = ["dc", "pdf", "pdfx", "xap", "xapMM", "xmp", "xmpMM"];
        return tags.reduce((acc, tag) => {
          const regex = new RegExp(`<${tag}:(.*?)>(.*?)</${tag}:(.*?)>`, 'gsm');
          const tagMatch = Array.from(xmp.matchAll(regex));
          if(tagMatch){
            return this.reduce_metadata(acc, tag, tagMatch);
          } else {
            return false;
          }
        }, {});
      } else {
        return 'error';
      }
    }
  }

  reduce_metadata(acc, type, list) {
    return list.reduce((acc, match) => {
      const [_, key, val, keyEnd] = match;
      let value = val
        .replace(/<[a-z]+:[a-z]+>/gi, " ")
        .replace(/<[a-z]+:.+["']>/gi, " ")
        .replace(/<\/[a-z]+:[a-z]+>/gi, " ")
        .replace(/<[a-z]+:.+\/>/gi, " ")
        .trim();

      const decodedValue = this.decode_value(value);

      if (decodedValue.ok) {
        acc[`${type},${key}`] = decodedValue.ok;
      }
      return acc;
    }, acc);
  }

  /**********************
  * infodata wrangling
  ***********************/

  parse_info_object(string) {
    const regexString = /\/([a-zA-Z]+)\s*\((.*?[\)]*[^\\])\)/g;
    const regexHex = /\/([^ /]+)\s*<(.*?)>/gs;
    const regexObjects = /\/([a-z]+)\s([0-9]+\s[0-9]+)\s.*?R/ig;
    const regexValues = /obj.*?\((.*?)\).*?endobj/s;

    const strings = this.get_typed_info_tuples(string, regexString);
    const hex = this.get_typed_info_tuples(string, regexHex);
    const objects = this.get_typed_info_tuples(string, regexHex, regexValues);

    const result = {};
    for (const [key, val] of [...strings, ...hex, ...objects]) {
      result[key] = this.fix_non_printable(this.fix_octal_utf16(val));
    }
    return result;
  }

  get_typed_info_tuples(string, regex, val_regex){
    const collection = [];
    let match;
    while ((match = regex.exec(string)) !== null) {
      const [, key, val] = match;
      if (val_regex){
        const obj = this.get_object(val);
        const valMatch = obj[0].match(val_regex);
        if (valMatch) {
          val = valMatch[1];
        }
      }
      const decodedValue = this.decode_value(val);
      if (decodedValue.ok) {
        collection.push([key, decodedValue.ok]);
      }
    }
    return collection;
  }

  /*************************
  * ref and object extractors
  **************************/

  encrypt_refs() {
    const regex = /\/Encrypt\s*[0-9].*?R/g;
    return this.extract_refs(regex);
  }

  info_refs() {
    const regex = /\/Info[\s0-9]*?R/g;
    return this.extract_refs(regex);
  }

  metadata_refs() {
    const regex = /\/Metadata[\s0-9]*?R/g;
    return this.extract_refs(regex);
  }

  extract_refs(regex){
    const matches = this.binary.match(regex);
    const uniqueMatches = [...new Set(matches)];
    return uniqueMatches;
  }

  get_object(objId) {
    if (typeof this.binary === 'string' && typeof objId === 'string' && objId.length <= 15) {
      objId = objId.replace(/ /g, "\\s");

      const regex = new RegExp(`[^0-9]${objId}.obj.*?endobj`, 'gs');
      return this.binary.match(regex);
    }
  }

  /****************
  * value decoding
  *****************/

  decode_value(value, hex = false) {
    if (value.startsWith("feff")) {
      let base16EncodedUtf16BigEndian = value.slice(4).replace(/[^0-9a-f]/gi, "");
      try {
        let bw = new BufferWrangler(base16EncodedUtf16BigEndian, 'hex');
        let utf16 = this.utf16_size_fix(bw.getBuffer());
        let string = bw.setBuffer(utf16).toString('utf16le');
        return { ok: string };
      } catch (error) {
        return 'error';
      }
    }

    if (value.startsWith("\xFE\xFF\\")) {
      let utf16Octal = value.slice(3);
      try {
        let string = this.fix_octal_utf16(utf16Octal);
        return { ok: string };
      } catch (error) {
        return 'error';
      }
    }

    if (value.startsWith("\xFE\xFF")) {
      let utf16 = value.slice(2);
      utf16 = this.utf16_size_fix(utf16);
      try {
        let bw = new BufferWrangler(utf16);
        let string = bw.toString('utf16le');
        return { ok: string };
      } catch (error) {
        return 'error';
      }
    }

    if (value.startsWith("\\376\\377")) {
      let rest = value.slice(6);
      try {
        let string = this.fix_octal_utf16(rest);
        return { ok: string };
      } catch (error) {
        return 'error';
      }
    }

    if (hex) {
      let cleanedValue = value.toLowerCase().replace(/[^0-9a-f]/gi, "");
      try {
        let bw = new BufferWrangler(cleanedValue, 'hex');
        let decoded = bw.getBuffer();
        return { ok: decoded };
      } catch (error) {
        return 'error';
      }
    }

    return { ok: value };
  }

//encoding fixes
  is_printable(val) {
    return /^[\x20-\x7E]*$/.test(val);
  }

  fix_non_printable(enc_str) {
    if (!this.is_printable(enc_str)) {
      let bw = new BufferWrangler(enc_str);
      return bw.toString();
    }
    return enc_str;
  }

  utf16_size_fix(enc_str) {
    if (enc_str.length % 2 === 1) {
      return BufferWrangler.concat([enc_str, Buffer.from([0])]);
    }
    return enc_str;
  }

  fix_octal_utf16(enc_str) {
    return enc_str.split(/\\[0-3][0-7][0-7]/)
      .map(part => this.do_fix_occtal_utf16(part))
      .join('');
  }

  do_fix_occtal_utf16(code) {
    if (code.startsWith("\\")) {
      let digits = code.slice(1).split('');
      let num = parseInt(digits[0]) * 64 + parseInt(digits[1]) * 8 + parseInt(digits[2]);
      return String.fromCharCode(num);
    }
    return code;
  }
}

class BufferWrangler {

  hexToDecimalLookupTable(){
	  return {0: 0, 1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8, 9: 9, a: 10, b: 11, c: 12, d: 13, e: 14, f: 15, A: 10, B: 11, C: 12, D: 13, E: 14, F: 15}
  };

  constructor(buffer, enc){
    this.setBuffer(buffer, enc);
  }

  setBuffer(buffer, enc){
    switch(enc){
      case "hex":
        if (buffer.length % 2 !== 0) {
            throw new Error('Invalid Hex string length.');
          }

          const resultLength = buffer.length / 2;
          const bytes = new Uint8Array(resultLength);

          for (let index = 0; index < resultLength; index++) {
            const highNibble = this.hexToDecimalLookupTable()[hexString[index * 2]];
            const lowNibble = this.hexToDecimalLookupTable()[hexString[(index * 2) + 1]];

            if (highNibble === undefined || lowNibble === undefined) {
	            throw new Error(`Invalid Hex character encountered at position ${index * 2}`);
            }

            bytes[index] = (highNibble << 4) | lowNibble;
          }
          this.buffer = bytes;
      case "utf16":
        if (buffer instanceof ArrayBuffer) {
          this.buffer = new Uint16Array(buffer);
        } else if (ArrayBuffer.isView(buffer)) {
          this.buffer = new Uint16Array(value.buffer, value.byteOffset, value.byteLength);
        } else {
          throw new TypeError(`Unsupported value, got \`${typeof value}\`.`);
        }
      default:
        if (buffer instanceof ArrayBuffer) {
          this.buffer = new Uint8Array(buffer);
        } else if (ArrayBuffer.isView(buffer)) {
          this.buffer = new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
        } else {
          throw new TypeError(`Unsupported value, got \`${typeof value}\`.`);
        }
    }
  }

  concat(arrays, totalLength) {
	  if (arrays.length === 0) {
		  return new Uint8Array(0);
	  }
    if(!totalLength){
      totalLength = arrays.reduce((accumulator, currentValue) => accumulator + currentValue.length, 0);
    }
	  
	  const returnValue = new Uint8Array(totalLength);

	  let offset = 0;
	  for (const array of arrays) {
		  this.assertUint8Array(array);
		  returnValue.set(array, offset);
		  offset += array.length;
	  }

	  return returnValue;
  }

  getBuffer(){
    return this.buffer;
  }

  toString(enc){
    const decoder = new TextDecoder();

    switch(enc){
      case 'utf16le':
        if(this.determine_endianness('big')){
          this.buffer = this.swapEndianess();
        }
        return decoder.decode(this.buffer);

      default:
        this.assertUint8Array();
        return decoder.decode(this.buffer);
    }
  }

  isUint8Array() {
    if (this.buffer.constructor === Uint8Array) {
	    return true;
    }
    return this.buffer.toString() === '[object Uint8Array]';
  }

  assertUint8Array() {
    if (!this.isUint8Array()) {
	    throw new TypeError(`Expected \`Uint8Array\`, got \`${typeof value}\``);
    }
  }

  swapEndianess() {
    return new Float64Array(new Int8Array(this.buffer.buffer).reverse().buffer).reverse();
  }

  determine_endianness(initialGuess) {
    let littleZeros = 0, bigZeros = 0;
    for (let i = 0; i < this.buffer.length; i += 2) {
      if (this.buffer[i] === 0) {
        if (initialGuess === 'big') {
          bigZeros++;
        } else {
          littleZeros++;
        }
      }
    }
    return bigZeros >= littleZeros ? 'big' : 'little';
  }
}