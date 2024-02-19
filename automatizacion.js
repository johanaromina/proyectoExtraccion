const fs = require('fs');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const pdfParse = require('pdf-parse');

async function extraerDatosDeFactura(pdfPath, datosFacturas) {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    const pdfData = await fs.promises.readFile(pdfPath);
    
    // Utilizar pdf-parse para extraer texto del PDF
    const { text } = await pdfParse(pdfData);

    // Aquí deberías escribir el código para extraer los datos específicos de la factura
    const alumnoMatch = text.match(/ALUMNO: (.+)$/m);
    const alumno = alumnoMatch ? alumnoMatch[1].trim() : 'No se encontró';
    
    const formaPagoMatch = text.match(/FORMA DE PAGO: (.+)$/m);
    const formaPago = formaPagoMatch ? formaPagoMatch[1].trim() : 'No se encontró';
    
    const dniMatch = text.match(/DNI: (\d+)/);
    const dni = dniMatch ? dniMatch[1] : 'No se encontró';

    const ImporteMatch = text.match(/Importe U.:(.+)$/m);
    const Importe = ImporteMatch ? ImporteMatch[1] : 'No se encontró';

    // Verificar si se capturaron correctamente los datos
    console.log('Alumno:', alumno);
    console.log('DNI:', dni);
    console.log('Forma de Pago:', formaPago);
    console.log('Importe U. :', Importe);

    // Agregar los datos extraídos a la matriz de datosFacturas
    datosFacturas.push({
        alumno,
        dni,
        formaPago,
        Importe
        // Agrega más campos según sea necesario
    });

    // Una vez extraídos los datos, cierra el navegador
    await browser.close();
}



// Función principal
async function main() {
    const facturasDirectory = 'D:\\Diseños en RPA\\proyecto extraccion\\facturas_pdf';
    const files = fs.readdirSync(facturasDirectory);

    const datosFacturas = [];

    for (const file of files) {
        if (file.endsWith('.pdf')) {
            const pdfPath = `${facturasDirectory}/${file}`;
            await extraerDatosDeFactura(pdfPath, datosFacturas);
        }
    }

    // Generar el reporte en Excel después de que se hayan extraído todos los datos de las facturas
    await generarReporteEnExcel(datosFacturas);

    console.log('Datos extraídos de los PDF:', datosFacturas);
}

// Función para generar un reporte en Excel con los datos de las facturas
async function generarReporteEnExcel(datosFacturas) {
    // Crea un nuevo libro de Excel
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Reporte Facturas');

    // Agrega encabezados a las columnas
    worksheet.columns = [
        { header: 'Alumno', key: 'alumno', width: 30 },
        { header: 'DNI', key: 'dni', width: 15 },
        { header: 'Forma de Pago', key: 'formaPago', width: 30 },
        { header: 'Importe U.', key: 'Importe U. ', width: 30 }
        // Agrega más encabezados según sea necesario
    ];

    // Agrega datos de las facturas al reporte
    datosFacturas.forEach(factura => {
        worksheet.addRow(factura);
    });

    // Guarda el reporte en un archivo Excel
    await workbook.xlsx.writeFile('./ReportesExcel/reporte_facturas.xlsx');
}

main().then(() => console.log('Proceso completado')).catch(console.error);
