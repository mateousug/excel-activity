const XLSX = require('xlsx');

// Crear datos de ejemplo (array de objetos)
const students = [
      { Nombre: 'Ana López', Edad: 20, Curso: 'Matemáticas' },
      { Nombre: 'Juan Pérez', Edad: 22, Curso: 'Historia' },
      { Nombre: 'María Gómez', Edad: 19, Curso: 'Programación' },
      { Nombre: 'Pablo Sanches', Edad: 20, Curso: 'Quimica' },
      { Nombre: 'María Gómez', Edad: 18, Curso: 'Gestión de empresas' },
];

// Convertir los datos en una hoja de cálculo
const worksheet = XLSX.utils.json_to_sheet(students);

// Crear un nuevo libro de trabajo (workbook)
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Estudiantes');

// Guardar el archivo Excel
XLSX.writeFile(workbook, 'output.xlsx');

console.log('Archivo Excel "output.xlsx" creado exitosamente.');

// Leer un archivo Excel existente
const inputFile = 'output.xlsx'; // Cambia a 'data/input.xlsx' si usas otro archivo
const workbookRead = XLSX.readFile(inputFile);

// Obtener la primera hoja del archivo
const sheetName = workbookRead.SheetNames[0];
const worksheetRead = workbookRead.Sheets[sheetName];

// Convertir la hoja a un array de objetos
const data = XLSX.utils.sheet_to_json(worksheetRead);

// Mostrar los datos en la consola
console.log('Contenido del archivo Excel:');
console.log(data);