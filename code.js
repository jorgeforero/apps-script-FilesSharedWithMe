/**
 * FilesSharedWithMe
 * Obtener los archivos compartidos conmigo en una hoja de cálculo y poder marcar archivos para
 * remover el permiso de edición
 */

/**
 * onOpen
 * Despliega el menú de con las opciones
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu( 'Opciones' )
      .addItem( '_ 🗄️ _ Obtener Archivos', 'getFilesSharedWithMe' )
      .addSeparator()
      .addItem( '_ ❌ _ Remover permiso', 'removeMeAsEditor' )
      .addToUi();
};

/**
 * getFilesSharedWithMe
 * Obtiene los archivos que hayan sido compartidos con la cuenta activa ( que ejecuta el script ) y registra
 * la información en una hoja de cálculo dada. La información que registra es: Name, Id, Type, Owner, Url y 
 * Edit ( booleano que indica si la cuenta activa tiene permisos de editor sobre el archivo )
 *   
 * @param {void} - void
 * @param {void} - void
 * @return {number} filesCounter - Número de archivos encontrados - Información en la hoja de cálculo dada ( si aplica )
 */
function getFilesSharedWithMe() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Limpia el contenido de la hoja
  sheet.getRange( 2, 1, sheet.getLastRow(), sheet.getLastColumn() ).clearContent();
  SpreadsheetApp.flush();
  // Obtiene el email del usuario activo
  let email = Session.getActiveUser().getEmail();
  // Contadores
  let filesCounter = 0;
  let filesInfo = [];
  // Busca los archivos compartido 
  let files = DriveApp.searchFiles( 'sharedWithMe' );
  while ( files.hasNext() ) {
    // Obtiene el archivo
    let file = files.next();
    // Determina si dentro de los editores el correo dado tiene permisos de edición
    let editors = file.getEditors();
    let meEdit = editors.some( v => v.getEmail() === email );
    // Obtiene el dueño del archivo y el correo correspondiente
    let owner = file.getOwner();
    let ownerEmail = ( owner !== null ) ? owner.getEmail() : ' --- ';
    // Guarda la información obtenida en el arreglo 
    filesInfo.push( [ '', file.getName(), file.getId(), file.getMimeType(), ownerEmail , file.getUrl(), meEdit ] );
    filesCounter++;
  };
  // Si se encontraron archivo compartidos con la cuenta, se registran en la hoja de cálculo
  if ( filesCounter > 0 ) {
    // Guarda los datos desde la fila 2 - Asume que la primera fila contiene los encabezados de las columnas
    sheet.getRange( 2, 1, filesInfo.length, filesInfo[ 0 ].length ).setValues( filesInfo );
  };
  // Retorna el número de archivo encontrados
  SpreadsheetApp.getActiveSpreadsheet().toast( `Se encontraron ${filesCounter} archivos`, 'Status', 4 );
  return filesCounter;
};

/**
 * removeMeAsEditor
 * A partir de la información obtenida por la función getFilesSharedWithMe en la hoja de calculo, se remueve
 * el permiso de editor de los archivos cuya columna RemoveMe y meEdit esten marcados en true.
 * La hoja calcula es actualizada en la columna RemoveMe indicando los archivos a los cuales le fue removido el permiso
 * 
 * @param {void} - void
 * @param {void} - void
 * @return {number} filesCounter - Número de archivos a los que se les removio el permiso de editor - Información actualizaza en la hoja ( si aplica )
 */
function removeMeAsEditor() {
  // Obtiene el email del usuario activo
  let email = Session.getActiveUser().getEmail();
  SpreadsheetApp.getActiveSpreadsheet().toast( `Working...`, 'Status', 4 );
  // Contadores
  let filesCounter = 0;
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Obtiene los datos de los archivos encontrados por la función getFilesSharedWithMe
  let filesInfo = sheet.getDataRange().getValues();
  // Obtiene el header - Primera fila del arreglo
  let header = filesInfo.shift();
  // Obtiene los valores de la columna identificada con el nombre RemoveMe
  let flags = getColumnValues( header, filesInfo, 'RemoveMe' );
  for ( let indx=0; indx<filesInfo.length; indx++ ) {
    let record = getRowAsObject( filesInfo[ indx ], header );
    // Si la columna removeme esta vacia y la columna meedit es true, se remueve el permiso de editor
    if ( ( record.removeme ) && ( record.meedit ) ) {
      let file = DriveApp.getFileById( record.id );
      file.removeEditor( email );
      filesCounter++;
      // Marca la celda del registro en proceso como removida
      flags[ indx ] = [ 'ReMoVeD' ];
    };
  };//for
  // Si hubo cambios, se actualiza la hoja de cálculo fuente
  if ( filesCounter > 0 ) {
    // Actualiza solo la columna RemoveMe para marcar los archivos a los cuales se les eliminó el permiso de edición
    let removemeColIndx = getColumnIndex( header, 'RemoveMe' ) + 1;
    sheet.getRange( 2, removemeColIndx, flags.length, 1 ).setValues( flags );
  };
  SpreadsheetApp.getActiveSpreadsheet().toast( `Se removieron ${filesCounter} permisos de edición`, 'Status', 4 );
  return filesCounter;
};

/**
 * getRowAsObject
 * Obtiene un objeto con los valores de la fila dada: RowData. Toma los nombres de las llaves del parámtero Header. Las llaves
 * son dadas en minusculas y los espacios reemplazados por _
 * @param {array} RowData - Arreglo con los datos de la fila de la hoja
 * @param {array} Header - Arreglo con los nombres del encabezado de la hoja
 * @return {object} obj - Objeto con los datos de la fila y las propiedades nombradas de acuerdo a Header
 */
 function getRowAsObject( RowData, Header ) {
  let obj = {};
  for ( let indx=0; indx<RowData.length; indx++ ) {
    obj[ Header[ indx ].toLowerCase().replace( /\s/g, '_' ) ] = RowData[ indx ];
  };//for
  return obj;
};

/**
 * getColumnValues
 * Obtiene todos los datos de la columna con nombre ColName
 * @param {string} ColName - Nombre de la columna de acuerdo a header (this.tbheader)
 * @return {array} - Arreglo con valores de la columna. Formato => [ [1],[2],[3] ]
 */
function getColumnValues( Header, Data, ColName ) {
  // Extrae la columna del arreglo Bidimensional a un arreglo lineal
  let colIndex = getColumnIndex( Header, ColName );
  return Data.map( function( value ) { return [ value[ colIndex ] ]; });
};

/**
 * getColumnIndex
 * Obtiene el indice (index-0) de la columna con el nombre Name
 * @param {string} Name - Nombre de la columna de acuerdo a el header
 * @return {integer} - Indice de Name en header o -1 sino lo encontró
 */
function getColumnIndex( Header, Name ) {
  return Header.indexOf( Name ); 
};
