// object holding the settings with the key as settings name
// and an object describing the setting as the respective value
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
 * Create two cards to control the add-on settings and
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

/**
 * Save settings changed by the user in the Settings card
 *
 * @param {Object} e an event object
 * @return {UniversalActionResponse}
 */
function onSaveSettings(e) {
  const card = createSettingsCard('Settings have been updated.');

  let navigation = CardService.newNavigation()
      .pushCard(card);
  let actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}

/**
 * Resets settings to the defaults
 *
 * @param {Object} e an event object
 * @return {UniversalActionResponse}
 */
function onResetSettings(e){
  let settings = {};
  for(const key in default_settings) {
    settings[key] = default_settings[key].value;
  }

  PropertiesService.getUserProperties().setProperties(settings);

  const card = createSettingsCard('Settings have been reset.');
  let navigation = CardService.newNavigation()
      .updateCard(card);
  let actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}

/**
 * Create and return a built settings card.
 * @param {string} a message to be displayed prominently on the card
 * @return {Card}
 */
function createSettingsCard(msg) {
  const props = PropertiesService.getUserProperties();
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
  let cs = CardService.newCardBuilder()
    .setName('about')
    .setHeader(CardService.newCardHeader().setTitle('About'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(`This add-on manages an eBook library. It currently has support for PDF and ePub books.
        Features include facilities to rename books based on metadata, building statistics on the library.`)
      )
    )

  return cs.build();
}

/**
 * Helper function to pull a setting identified by its name from the user properties
 * If the setting is not found in the user properties, the value from default settings is used
 * In addition the default value is stored in the user properties
 *
 * @param {string} name of the setting
 * @return {string} the value of the setting
 */
function getSetting(setting){
  const props = PropertiesService.getUserProperties();
  let ps = props.getProperty(setting);
  if (!ps){
    ps = default_settings[setting].value;
    props.setProperty(setting, ps);
  }
  return ps;
}

/**
 * Helper function to pull a setting identifying a folder from the user properties by its name
 * If not found in user properties, the value from the default settings is used instead and
 * stored in user properties. Using the value, a folder is identified and returned.
 *
 * @param {string} name of the setting
 * @return {Folder} a folder from the user drive
 */
function getFolderFromSetting(setting){
  const setting = getSetting(setting);
  let folder;
  if (setting.match(/^\d[\w_-]{25,}$/gmi)){
    folder = DriveApp.getFolderById(setting);
  } else {
    folder = DriveApp.getFoldersByName(setting).next();
    PropertiesService.getUserProperties().setProperty(setting, folder.getId());
  }
  return folder;
}

