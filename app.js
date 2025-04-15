const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const path = require("path");

// Rutas de archivos
const plantillabio = "informe_biomedicas.docx";
const plantillasoc = "informe_sociales.docx";
const plantillaing = "informe_ingenierias.docx";
const inputFile = "data_escuelas.json";
const inputFile_alumnos = "data_alumnos.json"; // Datos a reemplazar
const outputDir = "informes_escuelas"; // Carpeta de salida

// Asegurar que la carpeta de salida existe
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir);
}

// Cargar los datos desde el JSON
const jsonData = JSON.parse(fs.readFileSync(inputFile, "utf8"));
const jsonDataAlumnos = JSON.parse(fs.readFileSync(inputFile_alumnos, "utf8"));

// Función para formatear los datos
const formatData = (data) => {
  let alumnos_ = jsonDataAlumnos.filter((alumno) => {
    return alumno.escuela === data.escuela;
  });

  return {
    area: data.area,
    escuela: data.escuela,
    total: data.total,
    nsp_mat: data.nsp_mat,
    ap_mat: data.ap_mat,
    dp_mat: data.ds_mat,
    nsp_rv: data.nsp_rv,
    ap_rv: data.ap_rv,
    dp_rv: data.ds_rv,
    alumnos: alumnos_,
  };
};

// Generar documentos
jsonData.forEach((data, index) => {
  // Determinar la plantilla según el área
  let plantillaPath;
  switch (data.area) {
    case "Biomedicas":
      plantillaPath = plantillabio;
      break;
    case "Sociales":
      plantillaPath = plantillasoc;
      break;
    case "Ingenierias":
      plantillaPath = plantillaing;
      break;
    default:
      console.error(`❌ Área desconocida: ${data.area}`);
      return; // Salir de la iteración si el área no es válida
  }
  // Leer la plantilla en cada iteración para evitar el error
  const templateContent = fs.readFileSync(plantillaPath, "binary");
  const zip = new PizZip(templateContent);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  const formattedData = formatData(data); // Personalizar los datos
  doc.render(formattedData); // Cargar los datos en la plantilla

  // Generar el buffer del documento final
  const buffer = doc.getZip().generate({ type: "nodebuffer" });

  // Guardar el documento con un nombre único
  const fileName = path.join(
    outputDir,
    `informe de nivelacion Escuela profesional de ${formattedData.escuela}.docx`
  );
  fs.writeFileSync(fileName, buffer);

  console.log(`✅ Documento generado: ${fileName}`);
});

console.log("📄 Todos los documentos han sido creados exitosamente.");
