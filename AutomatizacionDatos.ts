function main(workbook: ExcelScript.Workbook) {
    // En base a texto en columna, dividir el texto para poder procesarlo mejor
    // Obtener la hoja
    let hoja = workbook.getWorksheet('YCO96'); // Se pone un nombre arbitrario
  
    // Se especifica la columna 
    let columna = "B";
    //let columnaBatch = "A";
  
    // Obtener el rango utilizado en la hoja de cálculo
    let rangoUsado = hoja.getUsedRange();
  
    // Obtener el número total de filas en el rango utilizado
    let totalFilas = rangoUsado.getRowCount();
  
    // Rango dinámico desde la primera celda (A1) hasta la última fila en el rango utilizado
    let rangoDinamico = hoja.getRange(`${columna}1:${columna}${totalFilas}`);
    //let rangoDinamicoBatch = hoja.getRange(`${columnaBatch}1:${columnaBatch}${totalFilas}`);
  
    // Leer los datos en el rango dinámico
    let datos = rangoDinamico.getValues();
    //let datosBatch = rangoDinamicoBatch.getValues();
  
    // Implementación por cada columna y buscarlo con expresiones regulares
    // Expresión regular para extraer la información necesaria
    let regex = /(\d{8})\s+([\d,.]+)\s+KG.*?([\d,.]+)\s+KG/;
    let regexBatch = /Batch Number: (\d+)/;
  
    // Crear un objeto para almacenar los datos extraídos
    let datosExtraidos: { [counter: number]: { article: string, reqWeight: string, actualWeight: string, batch: string, description?: string } } = {};
  
    // Contador para objeto
    let counter = 1;
    let batch: string;
    // Recorrer todas las filas y extraer los datos
    for (let i = 0; i < totalFilas; i++) {
      let fila = String(datos[i][0]);
      //let batch = String(datosBatch[i][0]);
  
      // Ejecutar la expresión regular en la fila de datos
      let resultado = regex.exec(fila);
      let resultadoBatch = regexBatch.exec(fila)
  
      if (resultadoBatch) {
        batch = resultadoBatch[1].substring(1);
        continue;
      }
      else if (resultado) {
        let article = resultado[1];
        let reqWeight = resultado[2];
        let actualWeight = resultado[3];
  
        // Se mete la descripción de la siguiente celda
        let description: string;
        let filaProx = String(datos[i + 1][0])
        description = filaProx.trim();
  
        // Agregar los datos extraídos al objeto
        datosExtraidos[counter] = { article, reqWeight, actualWeight, batch, description };
  
        // Incrementar el contador
        counter++;
      }
    }
  
    // Recorrer datosExtraidos para agregar separadores
    const separador = {
      article: "OtraOrfa",
      reqWeight: "OtraOrfa",
      actualWeight: "OtraOrfa",
      batch: "OtraOrfa",
      description: "OtraOrfa"
    };
    const datosExtraidosArray = Object.values(datosExtraidos);
    const datosExtraidosConSeparador = {};
    let counter2 = 1;
    for (let i = 0; i < datosExtraidosArray.length; i++) {
      datosExtraidosConSeparador[counter2++] = datosExtraidosArray[i];
  
      if (i < datosExtraidosArray.length - 1 && datosExtraidosArray[i].batch !== datosExtraidosArray[i + 1].batch) {
        datosExtraidosConSeparador[counter2++] = separador;
      }
    }
    console.log(`Datos extraidos con separador: `);
    console.log(datosExtraidosConSeparador)
  
    // Llenar lo demas
    llenarTabla(workbook, datosExtraidosConSeparador, totalFilas);
  
  }
  
  function llenarTabla(workbook: ExcelScript.Workbook, datosExtraidos: { [counter: number]: { article: string, reqWeight: string, actualWeight: string, batch: string, description?: string } }, totalFilasOrigen: number) {
    // Obtén una referencia a la hoja donde deseas escribir los datos
    let hojaDestino = workbook.getWorksheet('DATOS');
  
    // Obtener la tabla con formato
    let tablaPequenia = hojaDestino.getTable('Datos');
  
    // Verificar que la tabla existe
    if (!tablaPequenia) {
      console.log("No se encontró la tabla con el nombre 'Datos'.");
      return;
    }
  
    // Obtener los datos de la tabla
    let rangoTabla = tablaPequenia.getRangeBetweenHeaderAndTotal();
    let datosTabla = rangoTabla.getValues();
  
    // Quitar el cero de batch
    for (let i = 0; i < datosTabla.length; i++) {
      datosTabla[i][0] = String(datosTabla[i][0]).substring(1);
    }
  
    let totalFilas = tablaPequenia.getRowCount();
    let columnaInicioH = 8;
    let filaInicioH = 9;
    let batch: string;
    let totalQTY: string;
  
    // Define la fila y la columna de inicio donde escribir los datos
    let filaInicio = 11; // Cambia este valor según sea necesario
    let columnaInicio = 6; // Cambia este valor según sea necesario
  
    // Recorrer los datos de la tabla
    for (let i = 0; i < datosTabla.length && i <= totalFilas; i++) {
  
      // Poner en Tabla Datos2
      batch = String(datosTabla[i][0]);
      totalQTY = String(datosTabla[i][1]);
      datosTabla[i].push(columnaInicioH);
  
      // addColumn(columnaInicioH, "Description")
  
      hojaDestino.getCell(filaInicioH, columnaInicioH).setValue(batch);
      hojaDestino.getCell(filaInicioH, columnaInicioH + 1).setValue(totalQTY);
  
      columnaInicioH += 5;
    }
  
    console.log(`Datos header:`);
    console.log(datosTabla);
  
    // Recorre el objeto datosExtraidos y escribe datos en la tabla
    // let datosPrimeraOrfa = {};
    // let filaFinalPrimerosDatos = 0;
  
    // Agregar primeros datos
    // for (let indice in datosExtraidos) {
    //   if (datosExtraidos[indice].article == "OtraOrfa") {
    //     filaFinalPrimerosDatos = Number(indice);
    //     break
    //     // dataWriter(workbook, datosExtraidos, filaInicio, datosTabla, hojaDestino, (Number(indice) + 1));
    //   }
  
    //   let originalDescription = datosExtraidos[indice].description;
    //   let article = datosExtraidos[indice].article;
  
    //   // Crea una copia del objeto actual
    //   let objCopy = Object.assign({}, datosExtraidos[indice]);
  
    //   // Agrega el índice como atributo al objeto copiado
    //   objCopy['indice'] = filaInicio;
    //   objCopy['fila'] = filaFinalPrimerosDatos;
  
    //   // Asigna el objeto copiado a una nueva propiedad en datosPrimeraOrfa
    //   datosPrimeraOrfa[indice] = objCopy;
  
    //   // Escribe los datos en las celdas de la hoja destino
    //   hojaDestino.getCell(filaInicio, columnaInicio).setValue(article);
  
    //   // Busca en el diccionario
    //   let description = dictionary(workbook, article);
    //   if (description == undefined) {
    //     description = "No se encontró en el diccionario" + "," + originalDescription;
    //   }
    //   hojaDestino.getCell(filaInicio, columnaInicio + 1).setValue(description);
    //   filaInicio = filaInicio + 1;
    // }
  
    // console.log(`Primera Orfa: `);
    // console.log(datosPrimeraOrfa)
  
    // Agregar datos numericos
    let indiceAnterior = 0;
    let longitud = Object.keys(datosExtraidos).length;
    for (let indice in datosExtraidos) {
      if (datosExtraidos[indice].article == "OtraOrfa") {
        dataWriter(workbook, datosExtraidos, datosTabla, hojaDestino, indiceAnterior, (Number(indice) - 1));
        indiceAnterior = Number(indice);
      }
      else if (datosExtraidos[indice].article != "OtraOrfa" && Number(indice) == longitud) {
        dataWriter(workbook, datosExtraidos, datosTabla, hojaDestino, indiceAnterior, Number(indice));
      }
    }
  }
  
  
  function dictionary(workbook: ExcelScript.Workbook, article: string) {
    // Obtener la hoja
    let hoja = workbook.getWorksheet('Diccionario'); // Se pone un nombre arbitrario
  
    // Se especifica la columna 
    let columna = "A";
    let materialDescription = "B";
  
    // Obtener el rango utilizado en la hoja de cálculo
    let rangoUsado = hoja.getUsedRange();
  
    // Obtener el número total de filas en el rango utilizado
    let totalFilas = rangoUsado.getRowCount();
  
    // Rango dinámico desde la primera celda (A1) hasta la última fila en el rango utilizado
    let rangoDinamico = hoja.getRange(`${columna}2:${columna}${totalFilas}`);
    let rangoDinamicoMaterialDescription = hoja.getRange(`${materialDescription}2:${materialDescription}${totalFilas}`);
  
    // Leer los datos en el rango dinámico
    let datos = rangoDinamico.getValues();
    let datosMaterialDescription = rangoDinamicoMaterialDescription.getValues();
  
    for (let i = 0; i < totalFilas - 1; i++) {
      let material = String(datos[i][0]);
      let description = String(datosMaterialDescription[i][0]);
      if (article == material) {
  
        return description;
      }
  
    }
  
  }
  
  function dataWriter(workbook: ExcelScript.Workbook, datosExtraidos: { [counter: number]: { article: string, reqWeight: string, actualWeight: string, batch: string, description?: string } }, datosTabla: (string | number | boolean)[][], hojaDestino: ExcelScript.Worksheet, inicioAnterior: number, inicioIndice: number) {
    // Agregar datos restantes
    let keys = Object.keys(datosExtraidos); // Obtén las llaves del objeto
    let filaInicio = 11;
  
    // Comenzamos desde el indice indicado
    for (let i = inicioAnterior; i < inicioIndice; i++) {
      let indice = keys[i];
      let reqWeight: string = datosExtraidos[indice].reqWeight;
      let actualWeight: string = datosExtraidos[indice].actualWeight;
      let batchSAP: string = datosExtraidos[indice].batch;
      let article: string = datosExtraidos[indice].article;
      let descripcionOriginal: string = datosExtraidos[indice].description;
  
      // console.log(`BatchSAP: ${batchSAP}, reqWeight: ${reqWeight}, actualweight: ${actualWeight}`);
  
      let columna: string = "";
      let total_qty: string;
  
      // Data header Finder
      for (let i = 0; i < datosTabla.length; i++) {
        if (batchSAP == datosTabla[i][0]) {
          columna = String(datosTabla[i][3]);
          total_qty = String(datosTabla[i][1]);
        }
      }
      // console.log(`Columna: ${columna}`);
      // console.log(`REQ Weigth - ${reqWeight}`);
      if (reqWeight.length <= 5) {
        reqWeight = reqWeight.replace(',', '.');
      }
      else if (reqWeight.length <= 7) {
        reqWeight = reqWeight.replace(',', '.');
      }
      else if (reqWeight.length <= 8) {
        reqWeight = reqWeight.replace(',', '');
      }
      else {
        reqWeight = reqWeight.replace(',', '.');
      }
      if (actualWeight.length <= 5) {
        actualWeight = actualWeight.replace(',', '.');
      }
      else if (actualWeight.length <= 7) {
        actualWeight = actualWeight.replace(',', '.');
      }
      else if (actualWeight.length <= 8) {
        actualWeight = actualWeight.replace(',', '');
      }
      else {
        actualWeight = actualWeight.replace(',', '.');
      }
      // console.log(`REQ Weigth Convertido - ${reqWeight}`);
      // Escribir datos
      let regla3 = (parseFloat(reqWeight) * 100) / parseFloat(total_qty);
  
      // Buscar en diccionario
      let description = dictionary(workbook, article);
      if (description == undefined) {
        description = "No se encontró en el diccionario" + "," + descripcionOriginal;
      }
      hojaDestino.getCell(filaInicio, parseInt(columna) - 2).setValue(article);
      hojaDestino.getCell(filaInicio, parseInt(columna) - 1).setValue(description);
  
      hojaDestino.getCell(filaInicio, parseInt(columna)).setValue(regla3);
      hojaDestino.getCell(filaInicio, parseInt(columna) + 1).setValue(parseFloat(reqWeight));
      hojaDestino.getCell(filaInicio, parseInt(columna) + 2).setValue(parseFloat(actualWeight));
  
      filaInicio++;
    }
  }