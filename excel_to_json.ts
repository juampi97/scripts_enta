/*  Script utilizado leer el archivo de configuracion y generar el json que le debe pasar
    el flujo de nube al flujo de escritorio.
    
    El archivo de configuracion debe tener en un hoja (Nombre configurable), 
    una tabla (nombre configurable) con 3 columnas (Nombre, Valor, Tipo) y todas las variables
*/

function main(workbook: ExcelScript.Workbook): string {
  // Get all the worksheets in the workbook. 
  let sheets = workbook.getWorksheets();
  // Get a list of all the worksheet names.
  let worksheetsNames = sheets.map((sheet) => sheet.getName());
  // Creo array tablas
  let texts = [["Nombre", "Valor", "Tipo"]];
  // Recorro todas las hojas
  for(let i = 0; i < worksheetsNames.length; i++){
    let tables = workbook.getWorksheet(worksheetsNames[i]).getTables();
    // Obtengo las tablas de cada hoja y las concateno en un mismo array
    for (let j = 0; j < tables.length; j++) {
      let textos = tables[j].getRange().getTexts();
      textos.shift()
      texts = texts.concat(textos)
    }
  }
  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  let returnObjects2: TableData[] = [];
  let returnObjects3: TableData[] = [];
  let returnObjects4: string;
  let returnObjects5: string;
  if (texts.length > 0) {
    returnObjects = returnObjectFromValues(texts);
    returnObjects2 = eliminarDuplicados(returnObjects);
    returnObjects3 = ordenarPorNombre(returnObjects2);
    returnObjects4 = formatoSalidaCloud(returnObjects3);
    returnObjects5 = eliminarVacios(returnObjects4);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects4));
  return returnObjects5;
}

// This function converts a 2D array of values into a generic JSON object.
// In this case, we have defined the TableData object, but any similar interface would work.
function returnObjectFromValues(values: string[][]): TableData[] {
  let objectArray: TableData[] = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i];
      continue;
    }

    let object = {};
    for (let j = 0; j < values[i].length; j++) {
      object[objectKeys[j]] = values[i][j];
    }

    objectArray.push(object as TableData);
  }

  return objectArray;
}

//Eliminar duplicados del campo Nombre
function eliminarDuplicados(data: TableData[]): TableData[] {
  const nombresUnicos: { [key: string]: boolean } = {};
  const resultado: TableData[] = [];

  data.forEach(item => {
    if (!nombresUnicos[item.Nombre]) {
      nombresUnicos[item.Nombre] = true;
      resultado.push(item);
    }
  });

  return resultado;
}

//Ordenar por campo Nombre
function ordenarPorNombre(data: TableData[]): TableData[] {
  return data.slice().sort((a, b) => {
    if (a.Nombre < b.Nombre) return -1;
    if (a.Nombre > b.Nombre) return 1;
    return 0;
  });
}

//Generar string de salida para enviar a PA Desktop
function formatoSalidaCloud(data: TableData[]): string {
  let out_string: string = "{";
  data.forEach(item => {
    if (item.Tipo == "string") {
      out_string += `"${item.Nombre}":"${item.Valor}",`;
    } else if (item.Tipo == "int") {
      out_string += `"${item.Nombre}":${item.Valor},`;
    } else {
      out_string += `"${item.Nombre}":"${item.Valor}",`;
    }
  })
  out_string = out_string.slice(0, -1);
  out_string += "}";

  return out_string
}

//Filtrar datos vacios
function eliminarVacios(data: string): string {
  let out_string: string = "";

  out_string = data.replace('\"\":\"\",', '');

  return out_string
}


interface TableData {
  "Nombre": string;
  "Valor": string;
  "Tipo": string;
}